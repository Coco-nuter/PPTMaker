"""AI PPT Agent 的核心数据契约。"""

from pathlib import PurePosixPath, PureWindowsPath
from typing import Annotated, Literal, Self

from pydantic import (
    BaseModel,
    ConfigDict,
    Field,
    StringConstraints,
    field_validator,
    model_validator,
)

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
DisplayName = Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=255)]
RelativePath = Annotated[
    str, StringConstraints(strip_whitespace=True, min_length=1, max_length=260)
]
HexColor = Annotated[str, StringConstraints(pattern=r"^#[0-9A-Fa-f]{6}$")]
LanguageCode = Annotated[
    str, StringConstraints(strip_whitespace=True, pattern=r"^[a-z]{2}(?:-[A-Z]{2})?$")
]
LayoutName = Literal["cover", "bullets", "closing"]


class ContractModel(BaseModel):
    """所有领域合同共享的严格输入策略。"""

    model_config = ConfigDict(extra="forbid")


def _validate_relative_path(value: str) -> str:
    """只允许位于项目工作区内的规范相对路径。"""
    normalized = value.replace("\\", "/")
    posix_path = PurePosixPath(normalized)
    windows_path = PureWindowsPath(value)

    if posix_path.is_absolute() or windows_path.is_absolute():
        raise ValueError("path must be relative to the project workspace")
    if any(part in {".", ".."} for part in posix_path.parts):
        raise ValueError("path must not contain '.' or '..' segments")

    return normalized


class SourceRef(ContractModel):
    """用户材料在文稿中的稳定引用。"""

    source_id: StableId
    display_name: DisplayName
    source_type: Literal["text", "markdown", "pdf", "docx", "pptx", "xlsx"]
    relative_path: RelativePath

    _relative_path_must_be_safe = field_validator("relative_path")(_validate_relative_path)


class AssetRef(ContractModel):
    """可供页面使用的图片资源引用。"""

    asset_id: StableId
    display_name: DisplayName
    media_type: Literal["image/png", "image/jpeg"]
    relative_path: RelativePath
    alt_text: ShortText

    _relative_path_must_be_safe = field_validator("relative_path")(_validate_relative_path)


class ThemeSpec(ContractModel):
    """用于确定性渲染的基础主题。"""

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


class SlideSpec(ContractModel):
    """第一阶段支持的单页内容合同。"""

    slide_id: StableId = Field(frozen=True)
    layout: LayoutName
    title: Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=120)]
    subtitle: (
        Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=200)]
        | None
    ) = None
    bullets: list[ShortText] = Field(default_factory=list, max_length=6)
    source_ids: list[StableId] = Field(default_factory=list, max_length=20)
    asset_ids: list[StableId] = Field(default_factory=list, max_length=20)
    notes: Annotated[str, StringConstraints(strip_whitespace=True, max_length=2000)] | None = None

    @field_validator("source_ids", "asset_ids")
    @classmethod
    def references_must_be_unique(cls, value: list[str]) -> list[str]:
        """同一页面内不能重复引用同一对象。"""
        if len(value) != len(set(value)):
            raise ValueError("references on a slide must be unique")
        return value

    @model_validator(mode="after")
    def content_must_match_layout(self) -> Self:
        """为当前三种布局限定必需字段和容量。"""
        if self.layout == "bullets" and not self.bullets:
            raise ValueError("bullets layout requires at least one bullet")
        if self.layout != "bullets" and self.bullets:
            raise ValueError(f"{self.layout} layout does not accept bullets")
        return self


class DeckSpec(ContractModel):
    """演示文稿的唯一结构化事实源。"""

    schema_version: Literal["1.0"] = "1.0"
    deck_id: StableId
    title: Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=200)]
    audience: ShortText
    purpose: Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=500)]
    language: LanguageCode = "zh-CN"
    aspect_ratio: Literal["16:9"] = "16:9"
    theme: ThemeSpec
    sources: list[SourceRef] = Field(default_factory=list, max_length=50)
    assets: list[AssetRef] = Field(default_factory=list, max_length=50)
    slides: list[SlideSpec] = Field(min_length=1, max_length=30)

    @model_validator(mode="after")
    def ids_and_references_must_be_valid(self) -> Self:
        """保证 ID 全局唯一，并拒绝悬空引用。"""
        source_ids = [source.source_id for source in self.sources]
        asset_ids = [asset.asset_id for asset in self.assets]
        slide_ids = [slide.slide_id for slide in self.slides]

        if len(slide_ids) != len(set(slide_ids)):
            raise ValueError("slide_id values must be unique")
        if len(source_ids) != len(set(source_ids)):
            raise ValueError("source_id values must be unique")
        if len(asset_ids) != len(set(asset_ids)):
            raise ValueError("asset_id values must be unique")

        all_ids = [self.deck_id, *source_ids, *asset_ids, *slide_ids]
        if len(all_ids) != len(set(all_ids)):
            raise ValueError("all deck, source, asset, and slide IDs must be unique")

        known_sources = set(source_ids)
        known_assets = set(asset_ids)
        for slide in self.slides:
            unknown_sources = set(slide.source_ids) - known_sources
            if unknown_sources:
                raise ValueError(
                    f"slide {slide.slide_id} references unknown source IDs: "
                    f"{sorted(unknown_sources)}"
                )

            unknown_assets = set(slide.asset_ids) - known_assets
            if unknown_assets:
                raise ValueError(
                    f"slide {slide.slide_id} references unknown asset IDs: {sorted(unknown_assets)}"
                )

        return self
