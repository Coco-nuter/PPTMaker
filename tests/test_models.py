"""DeckSpec 数据契约测试。"""

import json
from pathlib import Path

import pytest
from pydantic import ValidationError

from models import DeckSpec

FIXTURE_PATH = Path(__file__).parent / "fixtures" / "sample_deck.json"


def load_sample_deck_data() -> dict[str, object]:
    """每次读取一个独立的示例字典，便于构造无效输入。"""
    return json.loads(FIXTURE_PATH.read_text(encoding="utf-8"))


def test_sample_deck_parses() -> None:
    deck = DeckSpec.model_validate_json(FIXTURE_PATH.read_text(encoding="utf-8"))

    assert deck.deck_id == "deck_sample"
    assert [slide.layout for slide in deck.slides] == ["cover", "bullets", "closing"]


def test_deck_round_trips_through_json() -> None:
    deck = DeckSpec.model_validate(load_sample_deck_data())

    serialized = deck.model_dump_json(indent=2)

    assert DeckSpec.model_validate_json(serialized) == deck


def test_duplicate_slide_id_is_rejected() -> None:
    data = load_sample_deck_data()
    slides = data["slides"]
    assert isinstance(slides, list)
    slides[1]["slide_id"] = slides[0]["slide_id"]

    with pytest.raises(ValidationError, match="slide_id values must be unique"):
        DeckSpec.model_validate(data)


def test_slide_id_cannot_be_changed_after_validation() -> None:
    deck = DeckSpec.model_validate(load_sample_deck_data())

    with pytest.raises(ValidationError, match="Field is frozen"):
        deck.slides[0].slide_id = "slide_replacement"


def test_invalid_layout_is_rejected() -> None:
    data = load_sample_deck_data()
    slides = data["slides"]
    assert isinstance(slides, list)
    slides[1]["layout"] = "section"

    with pytest.raises(ValidationError):
        DeckSpec.model_validate(data)


def test_invalid_theme_color_is_rejected() -> None:
    data = load_sample_deck_data()
    theme = data["theme"]
    assert isinstance(theme, dict)
    theme["primary_color"] = "navy"

    with pytest.raises(ValidationError):
        DeckSpec.model_validate(data)


@pytest.mark.parametrize("slide_count", [0, 31])
def test_invalid_page_count_is_rejected(slide_count: int) -> None:
    data = load_sample_deck_data()
    data["slides"] = [
        {
            "slide_id": f"slide_{index:03d}",
            "layout": "closing",
            "title": f"第 {index + 1} 页",
        }
        for index in range(slide_count)
    ]

    with pytest.raises(ValidationError):
        DeckSpec.model_validate(data)


def test_overlong_field_is_rejected() -> None:
    data = load_sample_deck_data()
    data["title"] = "x" * 201

    with pytest.raises(ValidationError):
        DeckSpec.model_validate(data)


def test_layout_content_mismatch_is_rejected() -> None:
    data = load_sample_deck_data()
    slides = data["slides"]
    assert isinstance(slides, list)
    slides[0]["bullets"] = ["封面不应包含要点列表"]

    with pytest.raises(ValidationError, match="cover layout does not accept bullets"):
        DeckSpec.model_validate(data)


def test_unknown_reference_is_rejected() -> None:
    data = load_sample_deck_data()
    slides = data["slides"]
    assert isinstance(slides, list)
    slides[1]["source_ids"] = ["source_missing"]

    with pytest.raises(ValidationError, match="unknown source IDs"):
        DeckSpec.model_validate(data)


@pytest.mark.parametrize("unsafe_path", ["C:/private/brief.md", "../brief.md"])
def test_unsafe_source_path_is_rejected(unsafe_path: str) -> None:
    data = load_sample_deck_data()
    sources = data["sources"]
    assert isinstance(sources, list)
    sources[0]["relative_path"] = unsafe_path

    with pytest.raises(ValidationError):
        DeckSpec.model_validate(data)
