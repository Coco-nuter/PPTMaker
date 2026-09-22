"""Streamlit 最小闭环 UI 测试。"""

import base64
from pathlib import Path
from unittest.mock import patch

import pytest
from streamlit.testing.v1 import AppTest

from demo_workflow import DemoArtifacts, DemoGenerationError

APP_PATH = Path(__file__).parents[1] / "app.py"
ONE_PIXEL_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
)


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


def test_app_displays_validated_sample_outline() -> None:
    app = AppTest.from_file(str(APP_PATH)).run(timeout=10)

    assert not app.exception
    assert app.session_state.stage == "idle"
    assert app.session_state.deck_spec.title == "AI PPT Agent 项目介绍"
    metrics = {metric.label: metric.value for metric in app.metric}
    assert metrics["受众"] == "项目评审人员"
    assert metrics["用途"] == "说明产品目标、核心能力和下一步安排"
    visible_markdown = "\n".join(element.value for element in app.markdown)
    assert "AI PPT Agent" in visible_markdown
    assert "核心能力" in visible_markdown
    assert "谢谢" in visible_markdown
    assert "通过自然语言描述演示目标" in visible_markdown
    visible_captions = "\n".join(element.value for element in app.caption)
    assert "布局：cover" in visible_captions
    assert "布局：bullets" in visible_captions
    assert "布局：closing" in visible_captions
    assert app.button[0].label == "生成 PPT"


def test_app_success_shows_all_previews_and_download(tmp_path: Path) -> None:
    artifacts = create_artifacts(tmp_path)
    app = AppTest.from_file(str(APP_PATH)).run(timeout=10)

    with patch("demo_workflow.generate_demo_project", return_value=artifacts):
        app.button[0].click().run(timeout=10)

    assert not app.exception
    assert app.session_state.stage == "ready"
    assert app.session_state.project_id == "project_test"
    assert len(app.session_state.preview_paths) == 3
    assert len(app.image) == 3
    assert len(app.download_button) == 1
    assert app.download_button[0].label == "下载 PPTX"


def test_app_failure_is_not_ready_and_has_no_download() -> None:
    app = AppTest.from_file(str(APP_PATH)).run(timeout=10)
    error = DemoGenerationError("project_failed", "PowerPoint is unavailable")

    with patch("demo_workflow.generate_demo_project", side_effect=error):
        app.button[0].click().run(timeout=10)

    assert not app.exception
    assert app.session_state.stage == "error"
    assert app.session_state.project_id == "project_failed"
    assert app.session_state.pptx_path is None
    assert app.session_state.preview_paths == []
    assert "PowerPoint is unavailable" in app.session_state.error_message
    assert len(app.download_button) == 0


@pytest.mark.integration
@pytest.mark.powerpoint
def test_app_real_powerpoint_closed_loop() -> None:
    app = AppTest.from_file(str(APP_PATH)).run(timeout=10)

    app.button[0].click().run(timeout=180)

    assert not app.exception
    assert app.session_state.stage == "ready"
    assert app.session_state.project_id.startswith("project_")
    assert Path(app.session_state.pptx_path).is_file()
    assert len(app.session_state.preview_paths) == 3
    assert all(Path(path).is_file() for path in app.session_state.preview_paths)
    assert len(app.image) == 3
    assert len(app.download_button) == 1
