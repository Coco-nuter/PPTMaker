"""AI PPT Agent 的无 LLM Streamlit 最小闭环。"""

import sys
from pathlib import Path

import streamlit as st

PROJECT_ROOT = Path(__file__).resolve().parent
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

import demo_workflow  # noqa: E402
from models import DeckSpec, SlideSpec  # noqa: E402

SAMPLE_DECK_PATH = PROJECT_ROOT / "tests" / "fixtures" / "sample_deck.json"
WORKSPACE_ROOT = PROJECT_ROOT / "workspace"
STAGES = {"idle", "rendering", "ready", "error"}


def _initialize_session() -> None:
    defaults = {
        "project_id": None,
        "deck_spec": None,
        "stage": "idle",
        "pptx_path": None,
        "preview_paths": [],
        "error_message": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

    if st.session_state.stage not in STAGES:
        st.session_state.stage = "idle"
    if st.session_state.deck_spec is None:
        st.session_state.deck_spec = demo_workflow.load_deck_spec(SAMPLE_DECK_PATH)


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


def _clear_artifacts_for_new_run() -> None:
    st.session_state.project_id = None
    st.session_state.pptx_path = None
    st.session_state.preview_paths = []
    st.session_state.error_message = None
    st.session_state.stage = "rendering"


def _generate(deck: DeckSpec) -> None:
    _clear_artifacts_for_new_run()
    with st.status("正在生成 PPTX 并调用 PowerPoint 渲染预览……", expanded=True) as status:
        try:
            artifacts = demo_workflow.generate_demo_project(deck, WORKSPACE_ROOT)
        except demo_workflow.DemoGenerationError as error:
            st.session_state.project_id = error.project_id
            st.session_state.stage = "error"
            st.session_state.error_message = error.reason
            status.update(label="生成失败", state="error")
            return
        except Exception as error:
            st.session_state.stage = "error"
            st.session_state.error_message = str(error)
            status.update(label="生成失败", state="error")
            return

        st.session_state.project_id = artifacts.project_id
        st.session_state.pptx_path = str(artifacts.pptx_path)
        st.session_state.preview_paths = [str(path) for path in artifacts.preview_paths]
        st.session_state.error_message = None
        st.session_state.stage = "ready"
        status.update(label="PPT 与真实预览已生成", state="complete", expanded=False)


def _show_ready_artifacts() -> None:
    pptx_path = Path(st.session_state.pptx_path)
    preview_paths = tuple(Path(path) for path in st.session_state.preview_paths)
    if (
        not pptx_path.is_file()
        or not preview_paths
        or any(not path.is_file() for path in preview_paths)
    ):
        st.session_state.stage = "error"
        st.session_state.pptx_path = None
        st.session_state.preview_paths = []
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


def main() -> None:
    st.set_page_config(page_title="AI PPT Agent Demo", page_icon="📊", layout="wide")
    _initialize_session()

    st.title("AI PPT Agent · 无 LLM Demo")
    st.caption("固定示例：DeckSpec 校验 → 可编辑 PPTX → PowerPoint COM 真实预览")
    deck: DeckSpec = st.session_state.deck_spec
    _show_outline(deck)

    if st.button("生成 PPT", type="primary", width="stretch"):
        _generate(deck)

    if st.session_state.stage == "rendering":
        st.info("正在生成，请稍候……")
    elif st.session_state.stage == "error":
        st.error(f"生成失败：{st.session_state.error_message}")
    elif st.session_state.stage == "ready":
        _show_ready_artifacts()


if __name__ == "__main__":
    main()
