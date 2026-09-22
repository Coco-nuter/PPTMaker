"""DeckSpec 规划器测试。"""

import json
from pathlib import Path
from typing import cast

import pytest

from llm import LLMInvalidOutputError
from models import DeckSpec
from planner import PlanningRequestError, PromptLoadError, load_plan_prompt, plan_deck

FIXTURE_PATH = Path(__file__).parent / "fixtures" / "basic_deck.json"


class FakeProvider:
    """记录规划调用并返回固定 DeckSpec。"""

    def __init__(self, deck: DeckSpec | dict[str, object]) -> None:
        self.deck = deck
        self.calls: list[tuple[str, str]] = []

    def generate_deck(self, user_request: str, system_prompt: str) -> DeckSpec:
        self.calls.append((user_request, system_prompt))
        return cast(DeckSpec, self.deck)


def load_sample_deck() -> DeckSpec:
    return DeckSpec.model_validate_json(FIXTURE_PATH.read_text(encoding="utf-8"))


def test_planner_returns_valid_deck_and_passes_prompt() -> None:
    provider = FakeProvider(load_sample_deck())

    deck = plan_deck(" 生成一份3页的项目介绍 ", provider=provider)

    assert isinstance(deck, DeckSpec)
    assert len(deck.slides) == 3
    assert provider.calls[0][0] == "生成一份3页的项目介绍"
    assert "每页只表达一个主要结论" in provider.calls[0][1]
    assert "不得编造" in provider.calls[0][1]
    assert "cover" in provider.calls[0][1]
    assert "slide_id" in provider.calls[0][1]
    assert "最多 6 个要点" in provider.calls[0][1]


@pytest.mark.parametrize("user_request", ["", "   ", "x" * 8_001])
def test_invalid_user_request_is_rejected_before_provider_call(user_request: str) -> None:
    provider = FakeProvider(load_sample_deck())

    with pytest.raises(PlanningRequestError):
        plan_deck(user_request, provider=provider)

    assert provider.calls == []


def test_explicit_page_count_must_match_model_output() -> None:
    provider = FakeProvider(load_sample_deck())

    with pytest.raises(LLMInvalidOutputError, match="明确要求 6 页"):
        plan_deck("生成一份6页的研究生开题汇报", provider=provider)


def test_unsupported_page_count_is_rejected_before_provider_call() -> None:
    provider = FakeProvider(load_sample_deck())

    with pytest.raises(PlanningRequestError, match="1 到 30"):
        plan_deck("生成一份31页的研究生开题汇报", provider=provider)

    assert provider.calls == []


def test_planner_revalidates_provider_result() -> None:
    invalid_deck = json.loads(FIXTURE_PATH.read_text(encoding="utf-8"))
    invalid_deck["slides"][1]["layout"] = "unsupported"
    provider = FakeProvider(invalid_deck)

    with pytest.raises(LLMInvalidOutputError, match="DeckSpec 不合法"):
        plan_deck("生成项目介绍", provider=provider)


def test_missing_or_empty_prompt_is_rejected(tmp_path: Path) -> None:
    with pytest.raises(PromptLoadError, match="无法读取"):
        load_plan_prompt(tmp_path / "missing.md")

    empty_prompt = tmp_path / "empty.md"
    empty_prompt.write_text("  ", encoding="utf-8")
    with pytest.raises(PromptLoadError, match="提示词为空"):
        load_plan_prompt(empty_prompt)
