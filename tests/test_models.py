"""DeckSpec 数据契约测试。"""

import json
from pathlib import Path

import pytest
from pydantic import ValidationError

from models import (
    BulletsSlideSpec,
    ClosingSlideSpec,
    ComparisonSlideSpec,
    CoverSlideSpec,
    DeckSpec,
    MetricsSlideSpec,
    ProcessSlideSpec,
    SectionSlideSpec,
    TimelineSlideSpec,
    TwoColumnSlideSpec,
)

FIXTURE_PATH = Path(__file__).parent / "fixtures" / "sample_deck.json"
EXPECTED_LAYOUTS = [
    "cover",
    "section",
    "bullets",
    "two_column",
    "metrics",
    "timeline",
    "process",
    "comparison",
    "closing",
]
EXPECTED_TYPES = [
    CoverSlideSpec,
    SectionSlideSpec,
    BulletsSlideSpec,
    TwoColumnSlideSpec,
    MetricsSlideSpec,
    TimelineSlideSpec,
    ProcessSlideSpec,
    ComparisonSlideSpec,
    ClosingSlideSpec,
]


def load_sample_deck_data() -> dict[str, object]:
    """每次读取一个独立的示例字典，便于构造无效输入。"""
    return json.loads(FIXTURE_PATH.read_text(encoding="utf-8"))


def slides_from(data: dict[str, object]) -> list[dict[str, object]]:
    slides = data["slides"]
    assert isinstance(slides, list)
    assert all(isinstance(slide, dict) for slide in slides)
    return slides


def test_sample_deck_parses_all_nine_discriminated_layouts() -> None:
    deck = DeckSpec.model_validate_json(FIXTURE_PATH.read_text(encoding="utf-8"))

    assert deck.schema_version == "2.0"
    assert deck.deck_id == "deck_sample"
    assert [slide.layout for slide in deck.slides] == EXPECTED_LAYOUTS
    assert [type(slide) for slide in deck.slides] == EXPECTED_TYPES


def test_slide_union_uses_layout_as_json_schema_discriminator() -> None:
    schema = DeckSpec.model_json_schema()
    slides_schema = schema["properties"]["slides"]
    discriminator = slides_schema["items"]["discriminator"]

    assert discriminator["propertyName"] == "layout"
    assert set(discriminator["mapping"]) == set(EXPECTED_LAYOUTS)


def test_deck_round_trips_through_json() -> None:
    deck = DeckSpec.model_validate(load_sample_deck_data())

    serialized = deck.model_dump_json(indent=2)

    assert DeckSpec.model_validate_json(serialized) == deck


def test_duplicate_slide_id_is_rejected() -> None:
    data = load_sample_deck_data()
    slides = slides_from(data)
    slides[1]["slide_id"] = slides[0]["slide_id"]

    with pytest.raises(ValidationError, match="slide_id values must be unique"):
        DeckSpec.model_validate(data)


def test_slide_id_cannot_be_changed_after_validation() -> None:
    deck = DeckSpec.model_validate(load_sample_deck_data())

    with pytest.raises(ValidationError, match="Field is frozen"):
        deck.slides[0].slide_id = "slide_replacement"


def test_invalid_layout_is_rejected() -> None:
    data = load_sample_deck_data()
    slides_from(data)[1]["layout"] = "chart"

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


@pytest.mark.parametrize(
    ("slide_index", "field", "value"),
    [
        (2, "bullets", []),
        (3, "left_items", []),
        (3, "right_items", ["item"] * 7),
        (4, "metrics", [{"value": "1", "label": "A", "description": "B"}]),
        (
            5,
            "items",
            [{"label": str(index), "title": "节点", "description": "说明"} for index in range(7)],
        ),
        (6, "steps", [{"title": "一步", "description": "说明"}]),
        (7, "rows", []),
    ],
)
def test_layout_collection_limits_are_enforced(
    slide_index: int,
    field: str,
    value: object,
) -> None:
    data = load_sample_deck_data()
    slides_from(data)[slide_index][field] = value

    with pytest.raises(ValidationError):
        DeckSpec.model_validate(data)


def test_section_requires_subtitle() -> None:
    data = load_sample_deck_data()
    slides_from(data)[1].pop("subtitle")

    with pytest.raises(ValidationError):
        DeckSpec.model_validate(data)


@pytest.mark.parametrize("field", ["sources", "assets"])
def test_removed_deck_fields_are_rejected(field: str) -> None:
    data = load_sample_deck_data()
    data[field] = []

    with pytest.raises(ValidationError, match="Extra inputs are not permitted"):
        DeckSpec.model_validate(data)


@pytest.mark.parametrize("field", ["source_ids", "asset_ids"])
def test_removed_slide_reference_fields_are_rejected(field: str) -> None:
    data = load_sample_deck_data()
    slides_from(data)[0][field] = []

    with pytest.raises(ValidationError, match="Extra inputs are not permitted"):
        DeckSpec.model_validate(data)


@pytest.mark.parametrize("field", ["x", "y", "primary_color", "drawing_code", "image_url"])
def test_model_cannot_supply_coordinates_colors_or_drawing_code(field: str) -> None:
    data = load_sample_deck_data()
    slides_from(data)[0][field] = "forbidden"

    with pytest.raises(ValidationError, match="Extra inputs are not permitted"):
        DeckSpec.model_validate(data)


def test_layout_specific_fields_do_not_leak_between_models() -> None:
    data = load_sample_deck_data()
    slides_from(data)[0]["bullets"] = ["封面不能包含要点字段"]

    with pytest.raises(ValidationError, match="Extra inputs are not permitted"):
        DeckSpec.model_validate(data)


def test_nested_visual_items_forbid_arbitrary_fields() -> None:
    data = load_sample_deck_data()
    metrics = slides_from(data)[4]["metrics"]
    assert isinstance(metrics, list)
    assert isinstance(metrics[0], dict)
    metrics[0]["fill_color"] = "#FFFFFF"

    with pytest.raises(ValidationError, match="Extra inputs are not permitted"):
        DeckSpec.model_validate(data)
