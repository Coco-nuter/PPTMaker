"""使用 Microsoft PowerPoint 生成 sample_deck 的真实 PNG 预览。"""

from pathlib import Path

from config import Settings
from models import DeckSpec
from pptx_renderer import render_deck
from preview import PreviewError, pptx_to_pngs

PROJECT_ROOT = Path(__file__).resolve().parents[1]
FIXTURE_PATH = PROJECT_ROOT / "tests" / "fixtures" / "sample_deck.json"
SOURCE_PPTX_PATH = PROJECT_ROOT / "output" / "sample_deck.pptx"
PREVIEW_ROOT = PROJECT_ROOT / "output" / "preview"


def main() -> None:
    settings = Settings()
    deck = DeckSpec.model_validate_json(FIXTURE_PATH.read_text(encoding="utf-8"))
    source_pptx = render_deck(deck, SOURCE_PPTX_PATH)

    try:
        result = pptx_to_pngs(
            source_pptx,
            PREVIEW_ROOT,
            timeout_seconds=settings.preview_timeout_seconds,
            width=settings.preview_width,
            height=settings.preview_height,
        )
    except PreviewError as error:
        raise SystemExit(f"Preview generation failed: {error}") from None

    print(f"PPTX: {result.source_pptx}")
    for png_path in result.png_paths:
        print(f"PNG:  {png_path}")


if __name__ == "__main__":
    main()
