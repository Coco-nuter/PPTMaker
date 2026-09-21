"""固定 DeckSpec 到可编辑 PPTX 的渲染测试。"""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE

from models import DeckSpec
from pptx_renderer import (
    HORIZONTAL_MARGIN,
    SLIDE_HEIGHT,
    SLIDE_WIDTH,
    VERTICAL_MARGIN,
    render_deck,
)

FIXTURE_PATH = Path(__file__).parent / "fixtures" / "sample_deck.json"


def load_sample_deck() -> DeckSpec:
    return DeckSpec.model_validate_json(FIXTURE_PATH.read_text(encoding="utf-8"))


def paragraph_texts(presentation: Presentation) -> list[str]:
    return [
        paragraph.text
        for slide in presentation.slides
        for shape in slide.shapes
        if shape.has_text_frame
        for paragraph in shape.text_frame.paragraphs
        if paragraph.text
    ]


def structural_summary(presentation: Presentation) -> list[tuple[object, ...]]:
    return [
        (
            shape.name,
            shape.shape_type,
            shape.left,
            shape.top,
            shape.width,
            shape.height,
            tuple(paragraph.text for paragraph in shape.text_frame.paragraphs)
            if shape.has_text_frame
            else (),
        )
        for slide in presentation.slides
        for shape in slide.shapes
    ]


def test_sample_deck_renders_and_reopens_without_mutation(tmp_path: Path) -> None:
    deck = load_sample_deck()
    original_json = deck.model_dump_json()
    output_path = tmp_path / "sample_deck.pptx"

    result = render_deck(deck, output_path)
    reopened = Presentation(result)

    assert result == output_path
    assert result.is_file()
    assert len(reopened.slides) == 3
    assert reopened.slide_width == SLIDE_WIDTH
    assert reopened.slide_height == SLIDE_HEIGHT
    assert deck.model_dump_json() == original_json
    assert all(
        shape.shape_type != MSO_SHAPE_TYPE.PICTURE
        for slide in reopened.slides
        for shape in slide.shapes
    )


def test_rendered_text_matches_sample_deck(tmp_path: Path) -> None:
    deck = load_sample_deck()
    presentation = Presentation(render_deck(deck, tmp_path / "content.pptx"))
    rendered_text = paragraph_texts(presentation)

    for slide in deck.slides:
        assert slide.title in rendered_text
        if slide.subtitle:
            assert slide.subtitle in rendered_text
        for bullet in slide.bullets:
            assert bullet in rendered_text


def test_text_uses_theme_fonts_and_shapes_stay_inside_safe_area(tmp_path: Path) -> None:
    deck = load_sample_deck()
    presentation = Presentation(render_deck(deck, tmp_path / "theme.pptx"))

    for slide in presentation.slides:
        for shape in slide.shapes:
            assert shape.left >= HORIZONTAL_MARGIN
            assert shape.left + shape.width <= SLIDE_WIDTH - HORIZONTAL_MARGIN
            assert shape.top >= VERTICAL_MARGIN
            assert shape.top + shape.height <= SLIDE_HEIGHT - VERTICAL_MARGIN

            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    expected_font = (
                        deck.theme.heading_font if "Title" in shape.name else deck.theme.body_font
                    )
                    assert run.font.name == expected_font

    assert presentation.slides[0].background.fill.fore_color.rgb == RGBColor.from_string(
        deck.theme.background_color.removeprefix("#")
    )
    assert presentation.slides[1].background.fill.fore_color.rgb == RGBColor.from_string(
        deck.theme.background_color.removeprefix("#")
    )
    assert presentation.slides[2].background.fill.fore_color.rgb == RGBColor.from_string(
        deck.theme.primary_color.removeprefix("#")
    )


def test_same_deck_produces_same_structure(tmp_path: Path) -> None:
    deck = load_sample_deck()
    first = Presentation(render_deck(deck, tmp_path / "first.pptx"))
    second = Presentation(render_deck(deck, tmp_path / "second.pptx"))

    assert structural_summary(first) == structural_summary(second)
