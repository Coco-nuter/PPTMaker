"""PowerPoint COM 真实预览链路测试。"""

import binascii
import json
import struct
import subprocess
import sys
import types
import zlib
from pathlib import Path
from typing import Any
from unittest.mock import MagicMock, patch

import pytest

from config import Settings
from models import DeckSpec
from pptx_renderer import render_deck
from preview import PreviewError, _powerpoint_worker, pptx_to_pngs

FIXTURE_PATH = Path(__file__).parent / "fixtures" / "sample_deck.json"


def render_sample_pptx(directory: Path) -> Path:
    deck = DeckSpec.model_validate_json(FIXTURE_PATH.read_text(encoding="utf-8"))
    return render_deck(deck, directory / "sample_deck.pptx")


def _png_chunk(kind: bytes, data: bytes) -> bytes:
    return (
        struct.pack(">I", len(data))
        + kind
        + data
        + struct.pack(">I", binascii.crc32(kind + data) & 0xFFFFFFFF)
    )


def create_png(path: Path, width: int = 1920, height: int = 1080) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    header = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    row = b"\x00" + (b"\x00\x00\x00" * width)
    data = zlib.compress(row * height)
    path.write_bytes(
        b"\x89PNG\r\n\x1a\n"
        + _png_chunk(b"IHDR", header)
        + _png_chunk(b"IDAT", data)
        + _png_chunk(b"IEND", b"")
    )


def successful_worker(page_count: int = 3, width: int = 1920, height: int = 1080):
    def run(command: list[str], **_: Any) -> subprocess.CompletedProcess[str]:
        output_directory = Path(command[command.index("--output-dir") + 1])
        manifest_path = Path(command[command.index("--manifest") + 1])
        for index in range(1, page_count + 1):
            create_png(output_directory / f"slide_{index:03d}.png", width, height)
        manifest_path.write_text(
            json.dumps({"status": "ok", "page_count": page_count}),
            encoding="utf-8",
        )
        return subprocess.CompletedProcess(command, 0, "", "")

    return run


def failed_worker(code: str, message: str):
    def run(command: list[str], **_: Any) -> subprocess.CompletedProcess[str]:
        manifest_path = Path(command[command.index("--manifest") + 1])
        manifest_path.write_text(
            json.dumps({"status": "error", "code": code, "message": message}),
            encoding="utf-8",
        )
        return subprocess.CompletedProcess(command, 1, "", "")

    return run


def test_settings_reads_powerpoint_preview_environment(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("PREVIEW_BACKEND", "powerpoint")
    monkeypatch.setenv("PREVIEW_TIMEOUT_SECONDS", "45")
    monkeypatch.setenv("PREVIEW_WIDTH", "1600")
    monkeypatch.setenv("PREVIEW_HEIGHT", "900")

    settings = Settings(_env_file=None)

    assert settings.preview_backend == "powerpoint"
    assert settings.preview_timeout_seconds == 45
    assert settings.preview_width == 1600
    assert settings.preview_height == 900


def test_non_windows_platform_is_rejected(tmp_path: Path) -> None:
    with (
        patch("preview.sys.platform", "linux"),
        pytest.raises(PreviewError, match="requires Windows"),
    ):
        pptx_to_pngs(tmp_path / "missing.pptx", tmp_path / "preview")


def test_missing_pptx_is_rejected_before_worker(tmp_path: Path) -> None:
    with (
        patch("preview.subprocess.run") as run_mock,
        pytest.raises(PreviewError, match="input file does not exist"),
    ):
        pptx_to_pngs(tmp_path / "missing.pptx", tmp_path / "preview")

    run_mock.assert_not_called()


def test_preview_uses_safe_subprocess_and_stable_names(tmp_path: Path) -> None:
    source_pptx = render_sample_pptx(tmp_path)

    with patch("preview.subprocess.run", side_effect=successful_worker()) as run_mock:
        result = pptx_to_pngs(
            source_pptx,
            tmp_path / "preview",
            timeout_seconds=17,
            width=1920,
            height=1080,
        )

    assert result.page_count == 3
    assert (result.width, result.height) == (1920, 1080)
    assert [path.name for path in result.png_paths] == [
        "slide_001.png",
        "slide_002.png",
        "slide_003.png",
    ]
    assert all(path.is_file() for path in result.png_paths)
    assert result.output_dir.name.startswith("preview-")
    assert not any(
        path.name.startswith(".preview-staging-") for path in result.output_dir.parent.iterdir()
    )

    command = run_mock.call_args.args[0]
    options = run_mock.call_args.kwargs
    assert isinstance(command, list)
    assert command[0] == sys.executable
    assert "--worker" in command
    assert options["shell"] is False
    assert options["capture_output"] is True
    assert options["timeout"] == 17


@pytest.mark.parametrize(
    ("code", "message", "expected"),
    [
        ("POWERPOINT_NOT_INSTALLED", "class not registered", "PowerPoint availability"),
        ("COM_START_FAILED", "automation error", "PowerPoint COM startup"),
        ("PPTX_OPEN_FAILED", "file is corrupt", "PowerPoint PPTX open"),
    ],
)
def test_worker_failures_have_clear_business_errors(
    tmp_path: Path,
    code: str,
    message: str,
    expected: str,
) -> None:
    source_pptx = render_sample_pptx(tmp_path)

    with (
        patch("preview.subprocess.run", side_effect=failed_worker(code, message)),
        pytest.raises(PreviewError, match=expected),
    ):
        pptx_to_pngs(source_pptx, tmp_path / "preview")


def test_timeout_reports_error_and_requests_powerpoint_cleanup(tmp_path: Path) -> None:
    source_pptx = render_sample_pptx(tmp_path)
    timeout = subprocess.TimeoutExpired(["python", "preview.py"], 3)

    with (
        patch("preview.subprocess.run", side_effect=timeout),
        patch("preview._terminate_powerpoint_process", return_value=None) as terminate_mock,
        pytest.raises(PreviewError, match="PowerPoint export timeout.*3 seconds"),
    ):
        pptx_to_pngs(source_pptx, tmp_path / "preview", timeout_seconds=3)

    terminate_mock.assert_called_once()


def test_page_count_mismatch_is_rejected(tmp_path: Path) -> None:
    source_pptx = render_sample_pptx(tmp_path)

    with (
        patch("preview.subprocess.run", side_effect=successful_worker(page_count=2)),
        pytest.raises(PreviewError, match="PPTX has 3 pages.*exported 2"),
    ):
        pptx_to_pngs(source_pptx, tmp_path / "preview")


def test_wrong_png_dimensions_are_rejected(tmp_path: Path) -> None:
    source_pptx = render_sample_pptx(tmp_path)

    with (
        patch(
            "preview.subprocess.run",
            side_effect=successful_worker(width=1280, height=720),
        ),
        pytest.raises(PreviewError, match="1280x720, expected 1920x1080"),
    ):
        pptx_to_pngs(source_pptx, tmp_path / "preview")


def test_failure_preserves_pptx_and_previous_success(tmp_path: Path) -> None:
    source_pptx = render_sample_pptx(tmp_path)
    original_bytes = source_pptx.read_bytes()
    preview_root = tmp_path / "preview"

    with patch("preview.subprocess.run", side_effect=successful_worker()):
        successful = pptx_to_pngs(source_pptx, preview_root)
    with (
        patch(
            "preview.subprocess.run",
            side_effect=failed_worker("SLIDE_EXPORT_FAILED", "slide 2 failed"),
        ),
        pytest.raises(PreviewError, match="PowerPoint slide export"),
    ):
        pptx_to_pngs(source_pptx, preview_root)

    assert successful.output_dir.is_dir()
    assert all(path.is_file() for path in successful.png_paths)
    assert source_pptx.read_bytes() == original_bytes
    assert len(list(preview_root.glob("preview-*"))) == 1
    assert not list(preview_root.glob(".preview-staging-*"))


def _fake_com_modules(application: MagicMock) -> dict[str, types.ModuleType]:
    pythoncom = types.ModuleType("pythoncom")
    pythoncom.CoInitialize = MagicMock()  # type: ignore[attr-defined]
    pythoncom.CoUninitialize = MagicMock()  # type: ignore[attr-defined]

    client = types.ModuleType("win32com.client")
    client.DispatchEx = MagicMock(return_value=application)  # type: ignore[attr-defined]
    win32com = types.ModuleType("win32com")
    win32com.client = client  # type: ignore[attr-defined]

    return {
        "pythoncom": pythoncom,
        "win32com": win32com,
        "win32com.client": client,
    }


def test_com_worker_initializes_exports_and_always_closes(tmp_path: Path) -> None:
    source_pptx = render_sample_pptx(tmp_path)
    presentation = MagicMock()
    presentation.Slides.Count = 3

    def slide_for(index: int) -> MagicMock:
        slide = MagicMock()
        slide.Export.side_effect = lambda path, *_: create_png(Path(path))
        return slide

    presentation.Slides.Item.side_effect = slide_for
    application = MagicMock()
    application.Presentations.Open.return_value = presentation
    modules = _fake_com_modules(application)
    output = tmp_path / "png"
    manifest = tmp_path / "result.json"
    pid_file = tmp_path / "powerpoint.pid"

    with (
        patch.dict(sys.modules, modules),
        patch("preview._powerpoint_process_ids", side_effect=[set(), {1234}]),
    ):
        exit_code = _powerpoint_worker(
            source_pptx,
            output,
            manifest,
            pid_file,
            1920,
            1080,
        )

    assert exit_code == 0
    modules["pythoncom"].CoInitialize.assert_called_once()  # type: ignore[attr-defined]
    modules["win32com.client"].DispatchEx.assert_called_once_with(  # type: ignore[attr-defined]
        "PowerPoint.Application"
    )
    application.Presentations.Open.assert_called_once_with(
        str(source_pptx),
        ReadOnly=True,
        Untitled=False,
        WithWindow=False,
    )
    assert presentation.Slides.Item.call_count == 3
    presentation.Close.assert_called_once()
    application.Quit.assert_called_once()
    modules["pythoncom"].CoUninitialize.assert_called_once()  # type: ignore[attr-defined]
    assert pid_file.read_text(encoding="ascii") == "1234"


def test_com_worker_closes_powerpoint_when_export_fails(tmp_path: Path) -> None:
    source_pptx = render_sample_pptx(tmp_path)
    presentation = MagicMock()
    presentation.Slides.Count = 1
    presentation.Slides.Item.return_value.Export.side_effect = RuntimeError("export failed")
    application = MagicMock()
    application.Presentations.Open.return_value = presentation
    modules = _fake_com_modules(application)

    with (
        patch.dict(sys.modules, modules),
        patch("preview._powerpoint_process_ids", side_effect=[set(), {1234}]),
    ):
        exit_code = _powerpoint_worker(
            source_pptx,
            tmp_path / "png",
            tmp_path / "result.json",
            tmp_path / "powerpoint.pid",
            1920,
            1080,
        )

    assert exit_code == 1
    presentation.Close.assert_called_once()
    application.Quit.assert_called_once()
    modules["pythoncom"].CoUninitialize.assert_called_once()  # type: ignore[attr-defined]


@pytest.mark.integration
@pytest.mark.powerpoint
def test_sample_deck_real_powerpoint_preview(tmp_path: Path) -> None:
    source_pptx = render_sample_pptx(tmp_path)
    try:
        result = pptx_to_pngs(source_pptx, tmp_path / "real-preview")
    except PreviewError as error:
        if "POWERPOINT_NOT_INSTALLED" in error.reason:
            pytest.skip(f"PowerPoint integration skipped: {error.reason}")
        raise

    assert result.page_count == 3
    assert len(result.png_paths) == 3
    assert all(path.is_file() for path in result.png_paths)
    assert (result.width, result.height) == (1920, 1080)
