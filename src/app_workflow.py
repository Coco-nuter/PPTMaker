"""Streamlit 创建流程使用的规划业务边界。"""

from __future__ import annotations

from dataclasses import dataclass

from llm import ModelProvider
from models import DeckSpec
from planner import plan_deck

MAX_INTAKE_TEXT_LENGTH = 2_000


class IntakeValidationError(ValueError):
    """需求输入缺失或超过当前 MVP 限制。"""


def _required_text(value: str, label: str, max_length: int = 200) -> str:
    normalized = value.strip()
    if not normalized:
        raise IntakeValidationError(f"{label}不能为空。")
    if len(normalized) > max_length:
        raise IntakeValidationError(f"{label}不能超过 {max_length} 个字符。")
    return normalized


@dataclass(frozen=True)
class DeckRequest:
    """第一阶段无素材规划所需的显式用户输入。"""

    requirements: str
    topic: str
    audience: str
    purpose: str
    slide_count: int

    def __post_init__(self) -> None:
        object.__setattr__(
            self,
            "requirements",
            _required_text(self.requirements, "自然语言需求", MAX_INTAKE_TEXT_LENGTH),
        )
        object.__setattr__(self, "topic", _required_text(self.topic, "PPT 主题"))
        object.__setattr__(self, "audience", _required_text(self.audience, "受众"))
        object.__setattr__(self, "purpose", _required_text(self.purpose, "用途", 500))
        if (
            isinstance(self.slide_count, bool)
            or not isinstance(self.slide_count, int)
            or not 1 <= self.slide_count <= 30
        ):
            raise IntakeValidationError("页数必须在 1 到 30 之间。")

    def to_prompt(self) -> str:
        """组合成包含明确页数与业务上下文的自然语言请求。"""
        return "\n".join(
            (
                f"生成一份{self.slide_count}页的PPT。",
                f"PPT 主题：{self.topic}",
                f"受众：{self.audience}",
                f"用途：{self.purpose}",
                f"具体要求：{self.requirements}",
            )
        )


def create_outline(request: DeckRequest, provider: ModelProvider | None = None) -> DeckSpec:
    """调用现有 planner，并只返回经过校验的候选大纲。"""
    return plan_deck(request.to_prompt(), provider=provider)
