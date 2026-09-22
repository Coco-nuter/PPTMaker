"""已确认 DeckSpec 的隔离项目生成与预览编排。"""

from __future__ import annotations

import uuid
from dataclasses import dataclass
from pathlib import Path

from models import DeckSpec
from pptx_renderer import render_deck
from preview import PreviewResult, pptx_to_pngs


class DemoGenerationError(RuntimeError):
    """携带失败项目 ID 的 Demo 生成异常。"""

    def __init__(self, project_id: str, reason: str) -> None:
        self.project_id = project_id
        self.reason = reason
        super().__init__(f"project {project_id} failed: {reason}")


@dataclass(frozen=True)
class DemoArtifacts:
    """一次成功 Demo 生成的全部可展示产物。"""

    project_id: str
    project_dir: Path
    deck_path: Path
    pptx_path: Path
    preview_paths: tuple[Path, ...]


def load_deck_spec(path: str | Path) -> DeckSpec:
    """读取 JSON 并执行完整 DeckSpec 校验。"""
    source = Path(path)
    if not source.is_file():
        raise FileNotFoundError(f"DeckSpec fixture does not exist: {source}")
    return DeckSpec.model_validate_json(source.read_text(encoding="utf-8"))


def _create_project_directory(workspace_root: str | Path) -> tuple[str, Path]:
    root = Path(workspace_root).resolve()
    root.mkdir(parents=True, exist_ok=True)

    for _ in range(10):
        project_id = f"project_{uuid.uuid4().hex}"
        project_dir = root / project_id
        try:
            project_dir.mkdir()
        except FileExistsError:
            continue
        return project_id, project_dir
    raise RuntimeError("could not allocate a unique project directory")


def _write_deck_snapshot(deck: DeckSpec, destination: Path) -> None:
    temporary = destination.with_suffix(".json.tmp")
    temporary.write_text(deck.model_dump_json(indent=2), encoding="utf-8")
    temporary.replace(destination)


def generate_demo_project(
    deck: DeckSpec,
    workspace_root: str | Path,
) -> DemoArtifacts:
    """为一次按钮点击创建隔离项目，并生成真实 PPTX 与 PowerPoint 预览。"""
    validated_deck = DeckSpec.model_validate(deck.model_dump(mode="json"))
    project_id, project_dir = _create_project_directory(workspace_root)
    deck_path = project_dir / "deck.json"
    output_dir = project_dir / "output"
    preview_dir = project_dir / "preview"
    pptx_path = output_dir / "sample_deck.pptx"

    try:
        output_dir.mkdir()
        preview_dir.mkdir()
        _write_deck_snapshot(validated_deck, deck_path)
        rendered_pptx = render_deck(validated_deck, pptx_path)
        preview: PreviewResult = pptx_to_pngs(rendered_pptx, preview_dir)
    except Exception as error:
        raise DemoGenerationError(project_id, str(error)) from error

    if preview.page_count != len(validated_deck.slides):
        raise DemoGenerationError(
            project_id,
            f"DeckSpec has {len(validated_deck.slides)} slides but preview has "
            f"{preview.page_count}",
        )

    return DemoArtifacts(
        project_id=project_id,
        project_dir=project_dir,
        deck_path=deck_path,
        pptx_path=rendered_pptx,
        preview_paths=preview.png_paths,
    )
