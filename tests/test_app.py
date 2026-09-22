"""自然语言规划、确认和生成闭环的 Streamlit AppTest。"""

import base64
from pathlib import Path
from unittest.mock import patch

import pytest
from streamlit.testing.v1 import AppTest

from demo_workflow import DemoArtifacts, DemoGenerationError
from llm import FakeModelProvider, LLMInvalidOutputError, LLMTimeoutError
from models import DeckSpec

APP_PATH = Path(__file__).parents[1] / "app.py"
FIXTURE_PATH = Path(__file__).parent / "fixtures" / "sample_deck.json"
ONE_PIXEL_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
)


def load_sample_deck() -> DeckSpec:
    return DeckSpec.model_validate_json(FIXTURE_PATH.read_text(encoding="utf-8"))


def create_artifacts(directory: Path) -> DemoArtifacts:
    project_dir = directory / "project_test"
    output_dir = project_dir / "output"
    png_dir = project_dir / "preview" / "preview-test" / "png"
    output_dir.mkdir(parents=True)
    png_dir.mkdir(parents=True)
    deck_path = project_dir / "deck.json"
    deck_path.write_text("{}", encoding="utf-8")
    pptx_path = output_dir / "sample_deck.pptx"
    pptx_path.write_bytes(b"mock pptx")
    preview_paths = tuple(png_dir / f"slide_{index:03d}.png" for index in range(1, 4))
    for path in preview_paths:
        path.write_bytes(ONE_PIXEL_PNG)
    return DemoArtifacts(
        project_id="project_test",
        project_dir=project_dir,
        deck_path=deck_path,
        pptx_path=pptx_path,
        preview_paths=preview_paths,
    )


def button(app: AppTest, label: str):
    return next(item for item in app.button if item.label == label)


def fill_intake(app: AppTest, *, slide_count: int = 3) -> None:
    app.text_area[0].set_value("简洁蓝白风，突出背景、方案和计划")
    app.text_input[0].set_value("研究生开题汇报")
    app.text_input[1].set_value("导师和答辩委员会")
    app.text_input[2].set_value("说明研究价值和实施方案")
    app.number_input[0].set_value(slide_count)


def plan_with_fake(app: AppTest, provider: FakeModelProvider) -> AppTest:
    app.session_state["_model_provider"] = provider
    fill_intake(app)
    return button(app, "生成大纲").click().run(timeout=10)


def test_app_starts_with_structured_intake() -> None:
    app = AppTest.from_file(str(APP_PATH)).run(timeout=10)

    assert not app.exception
    assert app.session_state.stage == "intake"
    assert [item.label for item in app.text_input] == ["PPT 主题", "受众", "用途"]
    assert app.text_area[0].label == "自然语言需求"
    assert app.number_input[0].label == "页数"
    assert button(app, "生成大纲")
    assert len(app.download_button) == 0


def test_fake_provider_plans_and_displays_outline_without_rendering() -> None:
    provider = FakeModelProvider(load_sample_deck())
    app = AppTest.from_file(str(APP_PATH)).run(timeout=10)

    with patch("demo_workflow.generate_demo_project") as generate:
        app = plan_with_fake(app, provider)

    assert not app.exception
    assert app.session_state.stage == "outline"
    assert app.session_state.deck_spec.title == "AI PPT Agent 项目介绍"
    assert app.session_state.pptx_path is None
    assert app.session_state.preview_paths == []
    assert len(provider.calls) == 1
    generate.assert_not_called()
    visible_markdown = "\n".join(element.value for element in app.markdown)
    assert "1. AI PPT Agent" in visible_markdown
    assert "通过自然语言描述演示目标" in visible_markdown
    visible_captions = "\n".join(element.value for element in app.caption)
    assert "slide_id：slide_cover" in visible_captions
    assert "布局：bullets" in visible_captions
    assert button(app, "确认大纲")
    assert button(app, "重新规划")
    assert button(app, "放弃当前大纲")


def test_confirmed_outline_generates_previews_and_download(tmp_path: Path) -> None:
    artifacts = create_artifacts(tmp_path)
    app = AppTest.from_file(str(APP_PATH)).run(timeout=10)
    app = plan_with_fake(app, FakeModelProvider(load_sample_deck()))

    with patch("demo_workflow.generate_demo_project", return_value=artifacts) as generate:
        app = button(app, "确认大纲").click().run(timeout=10)

    assert not app.exception
    assert app.session_state.stage == "ready"
    assert app.session_state.project_id == "project_test"
    assert app.session_state.successful_deck_spec == load_sample_deck()
    assert len(app.session_state.preview_paths) == 3
    assert len(app.image) == 3
    assert len(app.download_button) == 1
    assert app.download_button[0].label == "下载 PPTX"
    generate.assert_called_once()


def test_replan_replaces_candidate_and_abandon_returns_to_intake() -> None:
    first = load_sample_deck()
    second = first.model_copy(
        update={"title": "重新规划后的标题", "deck_id": "deck_replanned"}, deep=True
    )
    provider = FakeModelProvider([first, second])
    app = AppTest.from_file(str(APP_PATH)).run(timeout=10)
    app = plan_with_fake(app, provider)

    app = button(app, "重新规划").click().run(timeout=10)

    assert not app.exception
    assert app.session_state.stage == "outline"
    assert app.session_state.deck_spec.title == "重新规划后的标题"
    assert app.session_state.deck_spec.deck_id == "deck_replanned"
    assert len(provider.calls) == 2

    app = button(app, "放弃当前大纲").click().run(timeout=10)

    assert app.session_state.stage == "intake"
    assert app.session_state.deck_spec is None


@pytest.mark.parametrize(
    "planning_error",
    [LLMTimeoutError("provider timeout"), LLMInvalidOutputError("invalid DeckSpec")],
)
def test_planning_failure_keeps_last_successful_artifacts(
    tmp_path: Path, planning_error: Exception
) -> None:
    artifacts = create_artifacts(tmp_path)
    deck = load_sample_deck()
    app = AppTest.from_file(str(APP_PATH)).run(timeout=10)
    app = plan_with_fake(app, FakeModelProvider(deck))
    with patch("demo_workflow.generate_demo_project", return_value=artifacts):
        app = button(app, "确认大纲").click().run(timeout=10)

    app = button(app, "规划新 PPT").click().run(timeout=10)
    failing_provider = FakeModelProvider(deck, error=planning_error)
    app.session_state["_model_provider"] = failing_provider
    fill_intake(app)
    app = button(app, "生成大纲").click().run(timeout=10)

    assert not app.exception
    assert app.session_state.stage == "error"
    assert app.session_state.deck_spec is None
    assert app.session_state.project_id == "project_test"
    assert app.session_state.pptx_path == str(artifacts.pptx_path)
    assert app.session_state.preview_paths == [str(path) for path in artifacts.preview_paths]
    assert len(app.image) == 3
    assert len(app.download_button) == 1
    assert "本次失败未覆盖上一次成功结果" in "\n".join(item.value for item in app.caption)


def test_render_failure_does_not_offer_failed_artifact() -> None:
    app = AppTest.from_file(str(APP_PATH)).run(timeout=10)
    app = plan_with_fake(app, FakeModelProvider(load_sample_deck()))
    error = DemoGenerationError("project_failed", "PowerPoint is unavailable")

    with patch("demo_workflow.generate_demo_project", side_effect=error):
        app = button(app, "确认大纲").click().run(timeout=10)

    assert not app.exception
    assert app.session_state.stage == "error"
    assert app.session_state.failed_project_id == "project_failed"
    assert app.session_state.project_id is None
    assert app.session_state.pptx_path is None
    assert "PowerPoint is unavailable" in app.session_state.error_message
    assert len(app.download_button) == 0


@pytest.mark.integration
@pytest.mark.powerpoint
def test_app_real_powerpoint_closed_loop() -> None:
    app = AppTest.from_file(str(APP_PATH)).run(timeout=10)
    app = plan_with_fake(app, FakeModelProvider(load_sample_deck()))

    app = button(app, "确认大纲").click().run(timeout=180)

    assert not app.exception
    assert app.session_state.stage == "ready"
    assert app.session_state.project_id.startswith("project_")
    assert Path(app.session_state.pptx_path).is_file()
    assert len(app.session_state.preview_paths) == 3
    assert all(Path(path).is_file() for path in app.session_state.preview_paths)
    assert len(app.image) == 3
    assert len(app.download_button) == 1
