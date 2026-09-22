"""AI PPT Agent 的自然语言规划与确认生成闭环。"""

import sys
from pathlib import Path

import streamlit as st

PROJECT_ROOT = Path(__file__).resolve().parent
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

import app_workflow  # noqa: E402
import demo_workflow  # noqa: E402
from llm import LLMError, ModelProvider  # noqa: E402
from models import DeckSpec, SlideSpec  # noqa: E402
from planner import PlanningRequestError, PromptLoadError  # noqa: E402

WORKSPACE_ROOT = PROJECT_ROOT / "workspace"
STAGES = {"intake", "planning", "outline", "rendering", "ready", "error"}


def _initialize_session() -> None:
    defaults = {
        "project_id": None,
        "deck_spec": None,
        "successful_deck_spec": None,
        "planning_request": None,
        "stage": "intake",
        "pptx_path": None,
        "preview_paths": [],
        "error_message": None,
        "error_stage": None,
        "failed_project_id": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

    if st.session_state.stage not in STAGES:
        st.session_state.stage = "intake"


def _slide_content(slide: SlideSpec) -> str:
    if slide.bullets:
        return "\n".join(f"- {bullet}" for bullet in slide.bullets)
    return slide.subtitle or "（无补充内容）"


def _show_outline(deck: DeckSpec) -> None:
    st.header(deck.title)
    metadata = st.columns(2)
    metadata[0].metric("受众", deck.audience)
    metadata[1].metric("用途", deck.purpose)

    st.subheader("演示大纲")
    for index, slide in enumerate(deck.slides, start=1):
        with st.container(border=True):
            st.markdown(f"#### {index}. {slide.title}")
            st.caption(f"布局：{slide.layout} · slide_id：{slide.slide_id}")
            st.markdown(_slide_content(slide))


def _request_from_widgets() -> app_workflow.DeckRequest:
    request = app_workflow.DeckRequest(
        requirements=st.session_state.get("intake_requirements", ""),
        topic=st.session_state.get("intake_topic", ""),
        audience=st.session_state.get("intake_audience", ""),
        purpose=st.session_state.get("intake_purpose", ""),
        slide_count=int(st.session_state.get("intake_slide_count", 6)),
    )
    st.session_state.planning_request = request
    return request


def _provider_override() -> ModelProvider | None:
    """自动测试可注入 FakeProvider；正常运行不在 session 中保存真实客户端。"""
    return st.session_state.get("_model_provider")


def _plan_outline(*, reuse_request: bool = False) -> None:
    st.session_state.stage = "planning"
    st.session_state.deck_spec = None
    st.session_state.error_message = None
    st.session_state.error_stage = None
    st.session_state.failed_project_id = None

    try:
        if reuse_request:
            request = st.session_state.planning_request
            if not isinstance(request, app_workflow.DeckRequest):
                raise app_workflow.IntakeValidationError("原始需求丢失，请重新填写。")
        else:
            request = _request_from_widgets()
        deck = app_workflow.create_outline(request, provider=_provider_override())
    except (
        app_workflow.IntakeValidationError,
        LLMError,
        PlanningRequestError,
        PromptLoadError,
    ) as error:
        st.session_state.stage = "error"
        st.session_state.error_stage = "planning"
        st.session_state.error_message = str(error)
        return
    except Exception:
        st.session_state.stage = "error"
        st.session_state.error_stage = "planning"
        st.session_state.error_message = "规划发生未预期错误，请检查终端诊断。"
        return

    st.session_state.deck_spec = deck
    st.session_state.stage = "outline"


def _render_confirmed_outline() -> None:
    deck = st.session_state.deck_spec
    if not isinstance(deck, DeckSpec):
        st.session_state.stage = "error"
        st.session_state.error_stage = "rendering"
        st.session_state.error_message = "当前没有可确认的合法大纲，请重新规划。"
        return

    st.session_state.stage = "rendering"
    st.session_state.error_message = None
    st.session_state.error_stage = None
    st.session_state.failed_project_id = None

    try:
        artifacts = demo_workflow.generate_demo_project(deck, WORKSPACE_ROOT)
    except demo_workflow.DemoGenerationError as error:
        st.session_state.stage = "error"
        st.session_state.error_stage = "rendering"
        st.session_state.error_message = error.reason
        st.session_state.failed_project_id = error.project_id
        return
    except Exception:
        st.session_state.stage = "error"
        st.session_state.error_stage = "rendering"
        st.session_state.error_message = "生成发生未预期错误，请检查终端诊断。"
        return

    st.session_state.project_id = artifacts.project_id
    st.session_state.successful_deck_spec = deck
    st.session_state.pptx_path = str(artifacts.pptx_path)
    st.session_state.preview_paths = [str(path) for path in artifacts.preview_paths]
    st.session_state.error_message = None
    st.session_state.stage = "ready"


def _abandon_outline() -> None:
    st.session_state.deck_spec = None
    st.session_state.stage = "intake"
    st.session_state.error_message = None
    st.session_state.error_stage = None


def _show_intake() -> None:
    st.subheader("描述你想要的 PPT")
    saved_request = st.session_state.planning_request
    if isinstance(saved_request, app_workflow.DeckRequest):
        widget_defaults = {
            "intake_requirements": saved_request.requirements,
            "intake_topic": saved_request.topic,
            "intake_audience": saved_request.audience,
            "intake_purpose": saved_request.purpose,
            "intake_slide_count": saved_request.slide_count,
        }
        for key, value in widget_defaults.items():
            if key not in st.session_state:
                st.session_state[key] = value

    with st.form("deck_intake"):
        st.text_area(
            "自然语言需求",
            key="intake_requirements",
            placeholder="例如：简洁蓝白风，突出研究背景、方法、计划和预期成果",
            height=100,
        )
        st.text_input("PPT 主题", key="intake_topic", placeholder="例如：研究生开题汇报")
        st.text_input("受众", key="intake_audience", placeholder="例如：导师和答辩委员会")
        st.text_input("用途", key="intake_purpose", placeholder="例如：说明研究价值与实施方案")
        st.number_input(
            "页数",
            min_value=1,
            max_value=30,
            value=6,
            step=1,
            key="intake_slide_count",
        )
        submitted = st.form_submit_button("生成大纲", type="primary", width="stretch")

    if submitted:
        with st.spinner("正在规划故事线和页面大纲……"):
            _plan_outline()
        st.rerun()


def _show_outline_confirmation() -> None:
    deck = st.session_state.deck_spec
    if not isinstance(deck, DeckSpec):
        st.session_state.stage = "error"
        st.session_state.error_stage = "planning"
        st.session_state.error_message = "候选大纲丢失，请重新规划。"
        st.rerun()

    _show_outline(deck)
    confirm, replan, abandon = st.columns(3)
    if confirm.button("确认大纲", type="primary", width="stretch"):
        with st.status("正在生成 PPTX 并调用 PowerPoint 渲染预览……") as status:
            _render_confirmed_outline()
            if st.session_state.stage == "ready":
                status.update(label="PPT 与真实预览已生成", state="complete")
            else:
                status.update(label="生成失败", state="error")
        st.rerun()
    if replan.button("重新规划", width="stretch"):
        with st.spinner("正在重新规划，旧候选大纲已丢弃……"):
            _plan_outline(reuse_request=True)
        st.rerun()
    if abandon.button("放弃当前大纲", width="stretch"):
        _abandon_outline()
        st.rerun()


def _show_ready_artifacts() -> None:
    pptx_path = Path(st.session_state.pptx_path)
    preview_paths = tuple(Path(path) for path in st.session_state.preview_paths)
    if (
        not pptx_path.is_file()
        or not preview_paths
        or any(not path.is_file() for path in preview_paths)
    ):
        st.session_state.stage = "error"
        st.session_state.error_stage = "rendering"
        st.session_state.error_message = "当前项目产物缺失，请重新生成。"
        st.error(st.session_state.error_message)
        return

    st.success(f"生成成功 · 项目：{st.session_state.project_id}")
    st.subheader("PowerPoint 真实预览")
    for index, preview_path in enumerate(preview_paths, start=1):
        st.image(str(preview_path), caption=f"第 {index} 页", width="stretch")

    st.download_button(
        "下载 PPTX",
        data=pptx_path.read_bytes(),
        file_name=f"{st.session_state.project_id}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        type="primary",
        width="stretch",
    )


def _has_successful_artifacts() -> bool:
    return bool(
        st.session_state.project_id
        and st.session_state.pptx_path
        and st.session_state.preview_paths
    )


def _show_error() -> None:
    error_stage = "规划" if st.session_state.error_stage == "planning" else "生成"
    st.error(f"{error_stage}失败：{st.session_state.error_message}")

    actions = st.columns(2)
    if st.session_state.error_stage == "rendering" and isinstance(
        st.session_state.deck_spec, DeckSpec
    ):
        if actions[0].button("返回当前大纲", width="stretch"):
            st.session_state.stage = "outline"
            st.session_state.error_message = None
            st.rerun()
    if actions[1].button("返回需求输入", width="stretch"):
        _abandon_outline()
        st.rerun()

    if _has_successful_artifacts():
        st.divider()
        st.caption("本次失败未覆盖上一次成功结果。")
        _show_ready_artifacts()


def main() -> None:
    st.set_page_config(page_title="AI PPT Agent Demo", page_icon="📊", layout="wide")
    _initialize_session()

    st.title("AI PPT Agent")
    st.caption("自然语言需求 → DeckSpec 大纲确认 → 可编辑 PPTX → PowerPoint 真实预览")

    if st.session_state.stage == "intake":
        _show_intake()
    elif st.session_state.stage == "planning":
        st.info("正在规划大纲，请稍候……")
    elif st.session_state.stage == "outline":
        _show_outline_confirmation()
    elif st.session_state.stage == "rendering":
        st.info("正在生成 PPTX 和真实预览，请稍候……")
    elif st.session_state.stage == "ready":
        deck = st.session_state.successful_deck_spec
        if isinstance(deck, DeckSpec):
            _show_outline(deck)
        _show_ready_artifacts()
        if st.button("规划新 PPT", width="stretch"):
            _abandon_outline()
            st.rerun()
    elif st.session_state.stage == "error":
        _show_error()


if __name__ == "__main__":
    main()
