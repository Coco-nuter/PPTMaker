"""将固定的 DeckSpec 确定性渲染为可编辑 PPTX。"""

from collections.abc import Callable
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.slide import Slide
from pptx.util import Inches, Pt

from models import DeckSpec, SlideSpec, ThemeSpec

SLIDE_WIDTH = Inches(13.333333)
SLIDE_HEIGHT = Inches(7.5)
HORIZONTAL_MARGIN = Inches(0.8)
VERTICAL_MARGIN = Inches(0.55)
CONTENT_WIDTH = SLIDE_WIDTH - (2 * HORIZONTAL_MARGIN)


def _rgb(color: str) -> RGBColor:
    """将经过 DeckSpec 校验的 #RRGGBB 转成 PowerPoint 颜色。"""
    return RGBColor.from_string(color.removeprefix("#"))


def _set_background(slide: Slide, color: str) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = _rgb(color)


def _add_rectangle(
    slide: Slide,
    *,
    name: str,
    left: int,
    top: int,
    width: int,
    height: int,
    color: str,
) -> None:
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.name = name
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(color)
    shape.line.fill.background()


def _add_textbox(
    slide: Slide,
    *,
    name: str,
    text: str,
    left: int,
    top: int,
    width: int,
    height: int,
    font_name: str,
    font_size: int,
    color: str,
    bold: bool = False,
    alignment: PP_ALIGN = PP_ALIGN.LEFT,
    vertical_anchor: MSO_ANCHOR = MSO_ANCHOR.TOP,
) -> None:
    shape = slide.shapes.add_textbox(left, top, width, height)
    shape.name = name
    text_frame = shape.text_frame
    text_frame.clear()
    text_frame.word_wrap = True
    text_frame.margin_left = 0
    text_frame.margin_right = 0
    text_frame.margin_top = 0
    text_frame.margin_bottom = 0
    text_frame.vertical_anchor = vertical_anchor

    paragraph = text_frame.paragraphs[0]
    paragraph.alignment = alignment
    paragraph.line_spacing = 1.0
    run = paragraph.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = _rgb(color)


def _render_cover(slide: Slide, spec: SlideSpec, theme: ThemeSpec) -> None:
    _set_background(slide, theme.background_color)
    _add_rectangle(
        slide,
        name="Cover Accent",
        left=HORIZONTAL_MARGIN,
        top=Inches(1.25),
        width=Inches(0.12),
        height=Inches(4.45),
        color=theme.secondary_color,
    )
    _add_textbox(
        slide,
        name="Cover Title",
        text=spec.title,
        left=Inches(1.2),
        top=Inches(1.45),
        width=Inches(10.9),
        height=Inches(1.5),
        font_name=theme.heading_font,
        font_size=34,
        color=theme.primary_color,
        bold=True,
        vertical_anchor=MSO_ANCHOR.MIDDLE,
    )
    if spec.subtitle:
        _add_textbox(
            slide,
            name="Cover Subtitle",
            text=spec.subtitle,
            left=Inches(1.2),
            top=Inches(3.25),
            width=Inches(10.4),
            height=Inches(1.1),
            font_name=theme.body_font,
            font_size=20,
            color=theme.text_color,
            vertical_anchor=MSO_ANCHOR.MIDDLE,
        )


def _add_bullet_body(slide: Slide, spec: SlideSpec, theme: ThemeSpec, top: int) -> None:
    shape = slide.shapes.add_textbox(
        Inches(1.05),
        top,
        Inches(11.15),
        SLIDE_HEIGHT - top - VERTICAL_MARGIN,
    )
    shape.name = "Bullets Body"
    text_frame = shape.text_frame
    text_frame.clear()
    text_frame.word_wrap = True
    text_frame.margin_left = 0
    text_frame.margin_right = 0
    text_frame.margin_top = 0
    text_frame.margin_bottom = 0

    for index, bullet_text in enumerate(spec.bullets):
        paragraph = text_frame.paragraphs[0] if index == 0 else text_frame.add_paragraph()
        paragraph.level = 0
        paragraph.line_spacing = 1.1
        paragraph.space_after = Pt(14)
        run = paragraph.add_run()
        run.text = bullet_text
        run.font.name = theme.body_font
        run.font.size = Pt(20)
        run.font.color.rgb = _rgb(theme.text_color)

        paragraph_properties = paragraph._p.get_or_add_pPr()  # noqa: SLF001
        paragraph_properties.set("marL", str(Inches(0.34)))
        paragraph_properties.set("indent", str(-Inches(0.2)))
        bullet = paragraph_properties.makeelement(qn("a:buChar"), {"char": "•"})
        paragraph_properties.append(bullet)


def _render_bullets(slide: Slide, spec: SlideSpec, theme: ThemeSpec) -> None:
    _set_background(slide, theme.background_color)
    _add_rectangle(
        slide,
        name="Bullets Accent",
        left=HORIZONTAL_MARGIN,
        top=VERTICAL_MARGIN,
        width=Inches(0.08),
        height=Inches(0.78),
        color=theme.secondary_color,
    )
    _add_textbox(
        slide,
        name="Bullets Title",
        text=spec.title,
        left=Inches(1.05),
        top=VERTICAL_MARGIN,
        width=Inches(11.15),
        height=Inches(0.78),
        font_name=theme.heading_font,
        font_size=28,
        color=theme.primary_color,
        bold=True,
        vertical_anchor=MSO_ANCHOR.MIDDLE,
    )

    body_top = Inches(1.7)
    if spec.subtitle:
        _add_textbox(
            slide,
            name="Bullets Subtitle",
            text=spec.subtitle,
            left=Inches(1.05),
            top=Inches(1.42),
            width=Inches(11.15),
            height=Inches(0.5),
            font_name=theme.body_font,
            font_size=15,
            color=theme.secondary_color,
        )
        body_top = Inches(2.1)

    _add_bullet_body(slide, spec, theme, body_top)


def _render_closing(slide: Slide, spec: SlideSpec, theme: ThemeSpec) -> None:
    _set_background(slide, theme.primary_color)
    _add_textbox(
        slide,
        name="Closing Title",
        text=spec.title,
        left=HORIZONTAL_MARGIN,
        top=Inches(2.0),
        width=CONTENT_WIDTH,
        height=Inches(1.35),
        font_name=theme.heading_font,
        font_size=36,
        color=theme.background_color,
        bold=True,
        alignment=PP_ALIGN.CENTER,
        vertical_anchor=MSO_ANCHOR.MIDDLE,
    )
    _add_rectangle(
        slide,
        name="Closing Accent",
        left=Inches(5.55),
        top=Inches(3.55),
        width=Inches(2.23),
        height=Inches(0.08),
        color=theme.secondary_color,
    )
    if spec.subtitle:
        _add_textbox(
            slide,
            name="Closing Subtitle",
            text=spec.subtitle,
            left=HORIZONTAL_MARGIN,
            top=Inches(4.0),
            width=CONTENT_WIDTH,
            height=Inches(0.8),
            font_name=theme.body_font,
            font_size=18,
            color=theme.background_color,
            alignment=PP_ALIGN.CENTER,
            vertical_anchor=MSO_ANCHOR.MIDDLE,
        )


LayoutRenderer = Callable[[Slide, SlideSpec, ThemeSpec], None]
LAYOUT_RENDERERS: dict[str, LayoutRenderer] = {
    "cover": _render_cover,
    "bullets": _render_bullets,
    "closing": _render_closing,
}


def render_deck(deck: DeckSpec, output_path: str | Path) -> Path:
    """渲染 DeckSpec，保存后重新打开并核对页数。"""
    destination = Path(output_path)
    if destination.suffix.lower() != ".pptx":
        raise ValueError("output path must use the .pptx extension")
    destination.parent.mkdir(parents=True, exist_ok=True)

    presentation = Presentation()
    presentation.slide_width = SLIDE_WIDTH
    presentation.slide_height = SLIDE_HEIGHT
    presentation.core_properties.title = deck.title
    presentation.core_properties.subject = deck.purpose

    blank_layout = presentation.slide_layouts[6]
    for slide_spec in deck.slides:
        slide = presentation.slides.add_slide(blank_layout)
        LAYOUT_RENDERERS[slide_spec.layout](slide, slide_spec, deck.theme)

    presentation.save(destination)

    reopened = Presentation(destination)
    if len(reopened.slides) != len(deck.slides):
        raise RuntimeError("saved PPTX slide count does not match DeckSpec")

    return destination
