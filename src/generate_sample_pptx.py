"""从仓库 fixture 生成三页示例 PPTX。"""

from pathlib import Path

from models import DeckSpec
from pptx_renderer import render_deck

PROJECT_ROOT = Path(__file__).resolve().parents[1]
FIXTURE_PATH = PROJECT_ROOT / "tests" / "fixtures" / "sample_deck.json"
OUTPUT_PATH = PROJECT_ROOT / "output" / "sample_deck.pptx"


def main() -> None:
    deck = DeckSpec.model_validate_json(FIXTURE_PATH.read_text(encoding="utf-8"))
    destination = render_deck(deck, OUTPUT_PATH)
    print(f"Generated {len(deck.slides)} editable slides: {destination}")


if __name__ == "__main__":
    main()
