"""显式启用后调用真实 OpenAI API 的最小契约测试。"""

import os

import pytest

from config import Settings
from models import DeckSpec
from planner import plan_deck

pytestmark = pytest.mark.provider


def test_real_openai_generates_minimal_deck() -> None:
    if os.getenv("RUN_OPENAI_SMOKE_TESTS") != "1":
        pytest.skip("设置 RUN_OPENAI_SMOKE_TESTS=1 后才调用真实 OpenAI API")

    settings = Settings()
    if settings.openai_api_key is None or settings.openai_model is None:
        pytest.skip("真实 OpenAI smoke test 需要 OPENAI_API_KEY 和 OPENAI_MODEL")

    deck = plan_deck("生成一份3页的研究生开题汇报，简洁蓝白风", settings=settings)

    assert isinstance(deck, DeckSpec)
    assert len(deck.slides) == 3
    assert [slide.layout for slide in deck.slides] == ["cover", "bullets", "closing"]
