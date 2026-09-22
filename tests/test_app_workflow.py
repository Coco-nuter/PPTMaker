"""Streamlit 之外的需求输入与规划编排测试。"""

from pathlib import Path

import pytest

from app_workflow import DeckRequest, IntakeValidationError, create_outline
from llm import FakeModelProvider
from models import DeckSpec

FIXTURE_PATH = Path(__file__).parent / "fixtures" / "sample_deck.json"


def load_sample_deck() -> DeckSpec:
    return DeckSpec.model_validate_json(FIXTURE_PATH.read_text(encoding="utf-8"))


def test_deck_request_builds_explicit_planner_prompt() -> None:
    request = DeckRequest(
        requirements=" 简洁蓝白风 ",
        topic=" 研究生开题汇报 ",
        audience=" 导师和答辩委员会 ",
        purpose=" 说明研究价值 ",
        slide_count=3,
    )

    prompt = request.to_prompt()

    assert "生成一份3页的PPT" in prompt
    assert "PPT 主题：研究生开题汇报" in prompt
    assert "受众：导师和答辩委员会" in prompt
    assert "用途：说明研究价值" in prompt
    assert "具体要求：简洁蓝白风" in prompt


@pytest.mark.parametrize(
    ("field", "value"),
    [
        ("requirements", ""),
        ("topic", ""),
        ("audience", ""),
        ("purpose", ""),
        ("slide_count", 0),
        ("slide_count", 31),
        ("slide_count", 3.5),
    ],
)
def test_invalid_intake_is_rejected(field: str, value: object) -> None:
    values: dict[str, object] = {
        "requirements": "简洁蓝白风",
        "topic": "研究生开题汇报",
        "audience": "导师",
        "purpose": "开题答辩",
        "slide_count": 3,
    }
    values[field] = value

    with pytest.raises(IntakeValidationError):
        DeckRequest(**values)


def test_create_outline_uses_fake_provider_without_network() -> None:
    deck = load_sample_deck()
    provider = FakeModelProvider(deck)
    request = DeckRequest(
        requirements="简洁蓝白风",
        topic="研究生开题汇报",
        audience="导师",
        purpose="开题答辩",
        slide_count=3,
    )

    result = create_outline(request, provider=provider)

    assert result == deck
    assert result is not deck
    assert len(provider.calls) == 1
    assert "3页" in provider.calls[0][0]
