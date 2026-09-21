"""通过 Microsoft PowerPoint COM 生成最终 PPTX 的真实 PNG 预览。"""

from __future__ import annotations

import argparse
import csv
import gc
import json
import struct
import subprocess
import sys
import tempfile
import uuid
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from pptx import Presentation

from config import Settings

DEFAULT_TIMEOUT_SECONDS = 120.0
DEFAULT_PREVIEW_WIDTH = 1920
DEFAULT_PREVIEW_HEIGHT = 1080
POWERPOINT_PROG_ID = "PowerPoint.Application"
POWERPOINT_NOT_INSTALLED_HRESULTS = {-2147221164, -2147221005}


class PreviewError(RuntimeError):
    """预览链路中包含可操作原因的业务异常。"""

    def __init__(self, operation: str, reason: str) -> None:
        self.operation = operation
        self.reason = reason
        super().__init__(f"{operation} failed: {reason}")


@dataclass(frozen=True)
class PreviewResult:
    """一次独立 PowerPoint 预览转换的全部成功产物。"""

    source_pptx: Path
    output_dir: Path
    png_paths: tuple[Path, ...]
    page_count: int
    width: int
    height: int


def _require_windows() -> None:
    if sys.platform != "win32":
        raise PreviewError(
            "Platform check",
            "Microsoft PowerPoint preview requires Windows",
        )


def _require_positive(value: int | float, name: str) -> None:
    if value <= 0:
        raise PreviewError("Preview configuration", f"{name} must be greater than zero")


def _require_pptx(path: str | Path) -> Path:
    source = Path(path)
    if source.suffix.lower() != ".pptx":
        raise PreviewError("PPTX validation", f"expected a .pptx file: {source}")
    if not source.is_file():
        raise PreviewError("PPTX validation", f"input file does not exist: {source}")
    return source.resolve()


def _pptx_page_count(pptx_path: Path) -> int:
    try:
        return len(Presentation(pptx_path).slides)
    except Exception as error:
        raise PreviewError("PPTX validation", str(error)) from error


def _png_dimensions(path: Path) -> tuple[int, int]:
    try:
        with path.open("rb") as png_file:
            header = png_file.read(24)
    except OSError as error:
        raise PreviewError("PNG validation", f"cannot read {path}: {error}") from error
    if len(header) < 24 or header[:8] != b"\x89PNG\r\n\x1a\n" or header[12:16] != b"IHDR":
        raise PreviewError("PNG validation", f"PowerPoint produced an invalid PNG: {path}")
    return struct.unpack(">II", header[16:24])


def _create_final_directory_name(output_root: Path) -> Path:
    timestamp = datetime.now(UTC).strftime("%Y%m%dT%H%M%S%fZ")
    return output_root / f"preview-{timestamp}-{uuid.uuid4().hex[:8]}"


def _worker_command(
    source_pptx: Path,
    staging_directory: Path,
    manifest_path: Path,
    pid_path: Path,
    width: int,
    height: int,
) -> list[str]:
    return [
        sys.executable,
        str(Path(__file__).resolve()),
        "--worker",
        "--pptx",
        str(source_pptx),
        "--output-dir",
        str(staging_directory / "png"),
        "--manifest",
        str(manifest_path),
        "--pid-file",
        str(pid_path),
        "--width",
        str(width),
        "--height",
        str(height),
    ]


def _terminate_powerpoint_process(pid_path: Path) -> str | None:
    """超时时仅终止本次 COM 工作线程创建的 PowerPoint 进程。"""
    try:
        pid = int(pid_path.read_text(encoding="ascii").strip())
    except (OSError, ValueError):
        return None

    try:
        result = subprocess.run(
            ["taskkill", "/PID", str(pid), "/T", "/F"],
            check=False,
            shell=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=10,
        )
    except (OSError, subprocess.TimeoutExpired) as error:
        return f"could not terminate PowerPoint process {pid}: {error}"
    if result.returncode not in {0, 128}:
        details = result.stderr.strip() or result.stdout.strip() or "no process output"
        return f"could not terminate PowerPoint process {pid}: {details}"
    return None


def _run_worker(
    command: list[str],
    timeout_seconds: float,
    pid_path: Path,
) -> subprocess.CompletedProcess[str]:
    try:
        return subprocess.run(
            command,
            check=False,
            shell=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout_seconds,
        )
    except subprocess.TimeoutExpired as error:
        cleanup_error = _terminate_powerpoint_process(pid_path)
        details = f"after {timeout_seconds:g} seconds"
        if cleanup_error:
            details = f"{details}; {cleanup_error}"
        raise PreviewError("PowerPoint export timeout", details) from error
    except OSError as error:
        raise PreviewError("PowerPoint worker startup", str(error)) from error


def _read_worker_manifest(manifest_path: Path) -> dict[str, Any]:
    if not manifest_path.is_file():
        raise PreviewError(
            "PowerPoint worker",
            "worker exited without writing a result manifest",
        )
    try:
        data = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise PreviewError("PowerPoint worker", f"invalid result manifest: {error}") from error
    if not isinstance(data, dict):
        raise PreviewError("PowerPoint worker", "result manifest must be a JSON object")
    return data


def _raise_worker_error(
    manifest: dict[str, Any], process: subprocess.CompletedProcess[str]
) -> None:
    code = str(manifest.get("code", "COM_EXPORT_FAILED"))
    message = str(manifest.get("message", "unknown PowerPoint COM error"))
    operation_by_code = {
        "POWERPOINT_NOT_INSTALLED": "PowerPoint availability",
        "COM_START_FAILED": "PowerPoint COM startup",
        "PPTX_OPEN_FAILED": "PowerPoint PPTX open",
        "SLIDE_EXPORT_FAILED": "PowerPoint slide export",
        "PYWIN32_UNAVAILABLE": "PowerPoint COM startup",
    }
    operation = operation_by_code.get(code, "PowerPoint preview")
    diagnostics = process.stderr.strip() or process.stdout.strip()
    reason = f"{message} ({code})"
    if diagnostics:
        reason = f"{reason}; worker: {diagnostics}"
    raise PreviewError(operation, reason)


def _validate_staged_pngs(
    png_directory: Path,
    expected_count: int,
    width: int,
    height: int,
) -> tuple[Path, ...]:
    png_paths = tuple(sorted(png_directory.glob("slide_*.png")))
    expected_names = tuple(f"slide_{index:03d}.png" for index in range(1, expected_count + 1))
    actual_names = tuple(path.name for path in png_paths)
    if actual_names != expected_names:
        raise PreviewError(
            "Preview page validation",
            f"PPTX has {expected_count} pages but PNG files are {actual_names}",
        )
    for png_path in png_paths:
        dimensions = _png_dimensions(png_path)
        if dimensions != (width, height):
            raise PreviewError(
                "PNG dimension validation",
                f"{png_path.name} is {dimensions[0]}x{dimensions[1]}, expected {width}x{height}",
            )
    return png_paths


def pptx_to_pngs(
    pptx_path: str | Path,
    output_root: str | Path,
    *,
    timeout_seconds: float | None = None,
    width: int | None = None,
    height: int | None = None,
) -> PreviewResult:
    """在独立子进程中用 PowerPoint COM 逐页导出 PNG。"""
    _require_windows()
    settings = Settings()
    if settings.preview_backend != "powerpoint":
        raise PreviewError(
            "Preview configuration",
            f"unsupported PREVIEW_BACKEND: {settings.preview_backend}",
        )

    timeout = timeout_seconds if timeout_seconds is not None else settings.preview_timeout_seconds
    preview_width = width if width is not None else settings.preview_width
    preview_height = height if height is not None else settings.preview_height
    _require_positive(timeout, "PREVIEW_TIMEOUT_SECONDS")
    _require_positive(preview_width, "PREVIEW_WIDTH")
    _require_positive(preview_height, "PREVIEW_HEIGHT")

    source_pptx = _require_pptx(pptx_path)
    expected_page_count = _pptx_page_count(source_pptx)
    output_directory = Path(output_root).resolve()
    output_directory.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix=".preview-staging-", dir=output_directory) as temporary:
        staging_directory = Path(temporary)
        manifest_path = staging_directory / "result.json"
        pid_path = staging_directory / "powerpoint.pid"
        command = _worker_command(
            source_pptx,
            staging_directory,
            manifest_path,
            pid_path,
            preview_width,
            preview_height,
        )
        process = _run_worker(command, timeout, pid_path)
        manifest = _read_worker_manifest(manifest_path)

        if manifest.get("status") != "ok":
            _raise_worker_error(manifest, process)
        if process.returncode != 0:
            diagnostics = process.stderr.strip() or process.stdout.strip() or "no process output"
            raise PreviewError(
                "PowerPoint worker",
                f"exit code {process.returncode}; {diagnostics}",
            )

        reported_count = manifest.get("page_count")
        if reported_count != expected_page_count:
            raise PreviewError(
                "Preview page validation",
                f"PPTX has {expected_page_count} pages but PowerPoint exported {reported_count}",
            )
        _validate_staged_pngs(
            staging_directory / "png",
            expected_page_count,
            preview_width,
            preview_height,
        )

        final_directory = _create_final_directory_name(output_directory)
        staging_directory.replace(final_directory)

    png_paths = tuple(
        final_directory / "png" / f"slide_{index:03d}.png"
        for index in range(1, expected_page_count + 1)
    )
    return PreviewResult(
        source_pptx=source_pptx,
        output_dir=final_directory,
        png_paths=png_paths,
        page_count=expected_page_count,
        width=preview_width,
        height=preview_height,
    )


def _exception_hresult(error: Exception) -> int | None:
    value = getattr(error, "hresult", None)
    if isinstance(value, int):
        return value
    args = getattr(error, "args", ())
    if args and isinstance(args[0], int):
        return args[0]
    return None


def _write_manifest(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")


def _powerpoint_process_ids() -> set[int]:
    """读取当前用户可见的 PowerPoint 进程 ID，用于精确清理本次新实例。"""
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/FO", "CSV", "/NH"],
            check=False,
            shell=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=5,
        )
    except (OSError, subprocess.TimeoutExpired):
        return set()
    if result.returncode != 0:
        return set()

    process_ids: set[int] = set()
    for row in csv.reader(result.stdout.splitlines()):
        if len(row) < 2 or row[0].casefold() != "powerpnt.exe":
            continue
        try:
            process_ids.add(int(row[1]))
        except ValueError:
            continue
    return process_ids


def _powerpoint_worker(
    pptx_path: Path,
    output_directory: Path,
    manifest_path: Path,
    pid_path: Path,
    width: int,
    height: int,
) -> int:
    """COM 工作线程；必须在独立 Python 子进程中调用。"""
    try:
        import pythoncom
        import win32com.client
    except ImportError as error:
        _write_manifest(
            manifest_path,
            {"status": "error", "code": "PYWIN32_UNAVAILABLE", "message": str(error)},
        )
        return 1

    application = None
    presentation = None
    initialized = False
    try:
        pythoncom.CoInitialize()
        initialized = True
        existing_powerpoint_pids = _powerpoint_process_ids()
        try:
            application = win32com.client.DispatchEx(POWERPOINT_PROG_ID)
        except Exception as error:
            code = (
                "POWERPOINT_NOT_INSTALLED"
                if _exception_hresult(error) in POWERPOINT_NOT_INSTALLED_HRESULTS
                else "COM_START_FAILED"
            )
            _write_manifest(
                manifest_path,
                {"status": "error", "code": code, "message": str(error)},
            )
            return 1

        created_powerpoint_pids = _powerpoint_process_ids() - existing_powerpoint_pids
        if len(created_powerpoint_pids) == 1:
            powerpoint_pid = created_powerpoint_pids.pop()
            pid_path.write_text(str(powerpoint_pid), encoding="ascii")

        try:
            presentation = application.Presentations.Open(
                str(pptx_path),
                ReadOnly=True,
                Untitled=False,
                WithWindow=False,
            )
        except Exception as error:
            _write_manifest(
                manifest_path,
                {"status": "error", "code": "PPTX_OPEN_FAILED", "message": str(error)},
            )
            return 1

        output_directory.mkdir(parents=True, exist_ok=False)
        try:
            page_count = int(presentation.Slides.Count)
            for index in range(1, page_count + 1):
                png_path = output_directory / f"slide_{index:03d}.png"
                slide = presentation.Slides.Item(index)
                try:
                    slide.Export(str(png_path), "PNG", width, height)
                finally:
                    slide = None
        except Exception as error:
            _write_manifest(
                manifest_path,
                {"status": "error", "code": "SLIDE_EXPORT_FAILED", "message": str(error)},
            )
            return 1

        _write_manifest(manifest_path, {"status": "ok", "page_count": page_count})
        return 0
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
            presentation = None
        if application is not None:
            try:
                application.Quit()
            except Exception:
                pass
            application = None
        gc.collect()
        if initialized:
            pythoncom.CoUninitialize()


def _parse_worker_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--worker", action="store_true")
    parser.add_argument("--pptx", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--pid-file", type=Path, required=True)
    parser.add_argument("--width", type=int, required=True)
    parser.add_argument("--height", type=int, required=True)
    return parser.parse_args()


def _main() -> int:
    arguments = _parse_worker_arguments()
    if not arguments.worker:
        return 2
    return _powerpoint_worker(
        arguments.pptx,
        arguments.output_dir,
        arguments.manifest,
        arguments.pid_file,
        arguments.width,
        arguments.height,
    )


if __name__ == "__main__":
    raise SystemExit(_main())
