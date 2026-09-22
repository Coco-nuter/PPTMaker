"""自然语言演示文稿规划器。"""

from __future__ import annotations

import re
from pathlib import Path

from config import Settings
from llm import LLMInvalidOutputError, ModelProvider, OpenAIModelProvider, validate_deck_output
from models import DeckSpec

DEFAULT_PROMPT_PATH = Path(__file__).resolve().parents[1] / "prompts" / "plan_deck.md"
MAX_REQUEST_LENGTH = 8_000

_PAGE_COUNT_PATTERNS = (
    re.compile(r"(?<!第)(?P<count>\d{1,2})\s*页(?:的|PPT|幻灯片|演示|汇报)"),
    re.compile(r"(?P<count>\d{1,2})\s*(?:slides?|pages?)\b", re.IGNORECASE),
)


class PlanningRequestError(ValueError):
    """用户规划要求本身为空或超出当前限制。"""


class PromptLoadError(RuntimeError):
    """规划提示词不存在或为空。"""


def load_plan_prompt(path: str | Path = DEFAULT_PROMPT_PATH) -> str:
    """读取版本化的系统提示词。"""
    prompt_path = Path(path)
    try:
        prompt = prompt_path.read_text(encoding="utf-8").strip()
    except OSError as exc:
        raise PromptLoadError(f"无法读取规划提示词：{prompt_path}") from exc
    if not prompt:
        raise PromptLoadError(f"规划提示词为空：{prompt_path}")
    return prompt


def _normalize_request(user_request: str) -> str:
    request = user_request.strip()
    if not request:
        raise PlanningRequestError("PPT 要求不能为空。")
    if len(request) > MAX_REQUEST_LENGTH:
        raise PlanningRequestError(f"PPT 要求不能超过 {MAX_REQUEST_LENGTH} 个字符。")
    return request


def _requested_page_count(user_request: str) -> int | None:
    for pattern in _PAGE_COUNT_PATTERNS:
        match = pattern.search(user_request)
        if match is not None:
            count = int(match.group("count"))
            if not 1 <= count <= 30:
                raise PlanningRequestError("第一阶段 PPT 页数必须在 1 到 30 之间。")
            return count
    return None


def plan_deck(
    user_request: str,
    *,
    provider: ModelProvider | None = None,
    settings: Settings | None = None,
    prompt_path: str | Path = DEFAULT_PROMPT_PATH,
) -> DeckSpec:
    """将一条自然语言要求规划成经过二次校验的 DeckSpec。"""
    request = _normalize_request(user_request)
    requested_page_count = _requested_page_count(request)
    system_prompt = load_plan_prompt(prompt_path)
    active_provider = provider or OpenAIModelProvider(settings=settings)

    deck = validate_deck_output(active_provider.generate_deck(request, system_prompt))
    if requested_page_count is not None and len(deck.slides) != requested_page_count:
        raise LLMInvalidOutputError(
            f"模型返回 {len(deck.slides)} 页，但用户明确要求 {requested_page_count} 页。"
        )
    return deck
