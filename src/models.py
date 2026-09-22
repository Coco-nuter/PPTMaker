"""AI PPT Agent 的核心数据契约。"""

from typing import Annotated, Literal, Self

from pydantic import BaseModel, ConfigDict, Field, StringConstraints, model_validator

StableId = Annotated[
    str,
    StringConstraints(
        strip_whitespace=True,
        min_length=3,
        max_length=64,
        pattern=r"^[a-z][a-z0-9]*(?:_[a-z0-9]+)*$",
    ),
]
ShortText = Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=200)]
SlideTitle = Annotated[
    str, StringConstraints(strip_whitespace=True, min_length=1, max_length=120)
]
LabelText = Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=80)]
ValueText = Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=120)]
NotesText = Annotated[str, StringConstraints(strip_whitespace=True, max_length=2000)]
HexColor = Annotated[str, StringConstraints(pattern=r"^#[0-9A-Fa-f]{6}$")]
LanguageCode = Annotated[
    str, StringConstraints(strip_whitespace=True, pattern=r"^[a-z]{2}(?:-[A-Z]{2})?$")
]
LayoutName = Literal[
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


class ContractModel(BaseModel):
    """所有领域合同共享的严格输入策略。"""

    model_config = ConfigDict(extra="forbid")


class ThemeSpec(ContractModel):
    """统一控制所有页面的颜色和字体。"""

    name: Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=80)]
    primary_color: HexColor
    secondary_color: HexColor
    background_color: HexColor
    text_color: HexColor
    heading_font: Annotated[
        str, StringConstraints(strip_whitespace=True, min_length=1, max_length=100)
    ]
    body_font: Annotated[
        str, StringConstraints(strip_whitespace=True, min_length=1, max_length=100)
    ]


class BaseSlideSpec(ContractModel):
    """各布局共享的稳定身份和基础文字。"""

    slide_id: StableId = Field(frozen=True)
    title: SlideTitle
    notes: NotesText | None = None


class CoverSlideSpec(BaseSlideSpec):
    """封面页。"""

    layout: Literal["cover"]
    subtitle: ShortText | None = None


class SectionSlideSpec(BaseSlideSpec):
    """章节分隔页。"""

    layout: Literal["section"]
    subtitle: ShortText
    section_number: Annotated[
        str, StringConstraints(strip_whitespace=True, min_length=1, max_length=20)
    ] | None = None


class BulletsSlideSpec(BaseSlideSpec):
    """要点列表页。"""

    layout: Literal["bullets"]
    subtitle: ShortText | None = None
    bullets: list[ShortText] = Field(min_length=1, max_length=6)


class TwoColumnSlideSpec(BaseSlideSpec):
    """左右双栏页。"""

    layout: Literal["two_column"]
    left_title: LabelText
    left_items: list[ShortText] = Field(min_length=1, max_length=6)
    right_title: LabelText
    right_items: list[ShortText] = Field(min_length=1, max_length=6)


class MetricSpec(ContractModel):
    """指标卡中的一个可编辑指标。"""

    value: ValueText
    label: LabelText
    description: ShortText


class MetricsSlideSpec(BaseSlideSpec):
    """指标卡页面。"""

    layout: Literal["metrics"]
    metrics: list[MetricSpec] = Field(min_length=2, max_length=4)


class TimelineItemSpec(ContractModel):
    """时间轴中的一个节点。"""

    label: LabelText
    title: LabelText
    description: ShortText


class TimelineSlideSpec(BaseSlideSpec):
    """时间轴页面。"""

    layout: Literal["timeline"]
    items: list[TimelineItemSpec] = Field(min_length=2, max_length=6)


class ProcessStepSpec(ContractModel):
    """流程页中的一个步骤。"""

    title: LabelText
    description: ShortText


class ProcessSlideSpec(BaseSlideSpec):
    """流程步骤页面。"""

    layout: Literal["process"]
    steps: list[ProcessStepSpec] = Field(min_length=2, max_length=6)


class ComparisonRowSpec(ContractModel):
    """对比页中的一行。"""

    label: LabelText
    left_value: ValueText
    right_value: ValueText


class ComparisonSlideSpec(BaseSlideSpec):
    """左右对比页面。"""

    layout: Literal["comparison"]
    left_title: LabelText
    right_title: LabelText
    rows: list[ComparisonRowSpec] = Field(min_length=1, max_length=6)


class ClosingSlideSpec(BaseSlideSpec):
    """结束页。"""

    layout: Literal["closing"]
    subtitle: ShortText | None = None


SlideSpec = Annotated[
    CoverSlideSpec
    | SectionSlideSpec
    | BulletsSlideSpec
    | TwoColumnSlideSpec
    | MetricsSlideSpec
    | TimelineSlideSpec
    | ProcessSlideSpec
    | ComparisonSlideSpec
    | ClosingSlideSpec,
    Field(discriminator="layout"),
]


class DeckSpec(ContractModel):
    """演示文稿的唯一结构化事实源。"""

    schema_version: Literal["2.0"] = "2.0"
    deck_id: StableId
    title: Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=200)]
    audience: ShortText
    purpose: Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=500)]
    language: LanguageCode = "zh-CN"
    aspect_ratio: Literal["16:9"] = "16:9"
    theme: ThemeSpec
    slides: list[SlideSpec] = Field(min_length=1, max_length=30)

    @model_validator(mode="after")
    def slide_ids_must_be_unique(self) -> Self:
        """页面 ID 在整份文稿内唯一，且不能与 deck ID 冲突。"""
        slide_ids = [slide.slide_id for slide in self.slides]
        if len(slide_ids) != len(set(slide_ids)):
            raise ValueError("slide_id values must be unique")
        if self.deck_id in slide_ids:
            raise ValueError("deck_id and slide_id values must be unique")
        return self
