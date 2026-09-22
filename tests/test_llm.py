"""OpenAI 模型适配器的无网络单元测试。"""

import json
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import pytest
from openai import APIConnectionError, APITimeoutError

from config import Settings
from llm import (
    LLMEmptyOutputError,
    LLMInvalidOutputError,
    LLMNetworkError,
    LLMRefusalError,
    LLMTimeoutError,
    MissingAPIKeyError,
    OpenAIModelProvider,
)
from models import DeckSpec

FIXTURE_PATH = Path(__file__).parent / "fixtures" / "sample_deck.json"


def load_sample_deck() -> DeckSpec:
    return DeckSpec.model_validate_json(FIXTURE_PATH.read_text(encoding="utf-8"))


def make_settings(**overrides: object) -> Settings:
    values: dict[str, object] = {
        "openai_api_key": "test-api-key",
        "openai_model": "test-structured-model",
        "llm_timeout_seconds": 12,
        "_env_file": None,
    }
    values.update(overrides)
    return Settings(**values)


def make_client_response(parsed: object, output: list[object] | None = None) -> object:
    return SimpleNamespace(output_parsed=parsed, output=output or [])


def test_settings_reads_openai_environment(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("OPENAI_API_KEY", "environment-secret")
    monkeypatch.setenv("OPENAI_BASE_URL", "https://example.test/v1")
    monkeypatch.setenv("OPENAI_MODEL", "environment-model")
    monkeypatch.setenv("LLM_TIMEOUT_SECONDS", "45")

    settings = Settings(_env_file=None)

    assert settings.openai_api_key is not None
    assert settings.openai_api_key.get_secret_value() == "environment-secret"
    assert "environment-secret" not in repr(settings)
    assert settings.openai_base_url == "https://example.test/v1"
    assert settings.openai_model == "environment-model"
    assert settings.llm_timeout_seconds == 45


def test_missing_api_key_is_rejected_before_client_creation() -> None:
    settings = make_settings(openai_api_key=None)

    with patch("llm.OpenAI") as openai_constructor:
        with pytest.raises(MissingAPIKeyError, match="OPENAI_API_KEY"):
            OpenAIModelProvider(settings)

    openai_constructor.assert_not_called()


def test_provider_builds_client_from_configuration_without_exposing_key() -> None:
    settings = make_settings(openai_base_url="https://example.test/v1")
    parsed = load_sample_deck()

    with patch("llm.OpenAI") as openai_constructor:
        openai_constructor.return_value.responses.parse.return_value = make_client_response(parsed)
        provider = OpenAIModelProvider(settings)
        result = provider.generate_deck("生成三页项目介绍", "system rules")

    assert result == parsed
    openai_constructor.assert_called_once_with(
        api_key="test-api-key",
        base_url="https://example.test/v1",
        timeout=12.0,
        max_retries=1,
    )


def test_valid_structured_output_is_revalidated_as_deck_spec() -> None:
    parsed = load_sample_deck()
    client = MagicMock()
    client.responses.parse.return_value = make_client_response(parsed)
    provider = OpenAIModelProvider(make_settings(), client=client)

    result = provider.generate_deck("生成三页项目介绍", "system rules")

    assert result == parsed
    assert result is not parsed
    client.responses.parse.assert_called_once_with(
        model="test-structured-model",
        input=[
            {"role": "system", "content": "system rules"},
            {"role": "user", "content": "生成三页项目介绍"},
        ],
        text_format=DeckSpec,
        store=False,
        timeout=12.0,
    )


def test_invalid_structured_output_is_rejected() -> None:
    invalid = json.loads(FIXTURE_PATH.read_text(encoding="utf-8"))
    invalid["slides"][1]["layout"] = "unsupported"
    client = MagicMock()
    client.responses.parse.return_value = make_client_response(invalid)
    provider = OpenAIModelProvider(make_settings(), client=client)

    with pytest.raises(LLMInvalidOutputError, match="DeckSpec 不合法"):
        provider.generate_deck("生成三页项目介绍", "system rules")


def test_timeout_is_mapped_to_business_error() -> None:
    client = MagicMock()
    client.responses.parse.side_effect = APITimeoutError(request=MagicMock())
    provider = OpenAIModelProvider(make_settings(), client=client)

    with pytest.raises(LLMTimeoutError, match="12 秒"):
        provider.generate_deck("生成三页项目介绍", "system rules")


def test_network_failure_is_mapped_to_business_error() -> None:
    client = MagicMock()
    client.responses.parse.side_effect = APIConnectionError(request=MagicMock())
    provider = OpenAIModelProvider(make_settings(), client=client)

    with pytest.raises(LLMNetworkError, match="网络"):
        provider.generate_deck("生成三页项目介绍", "system rules")


def test_refusal_is_mapped_to_business_error() -> None:
    refusal = SimpleNamespace(type="refusal", refusal="该请求无法完成")
    output = [SimpleNamespace(content=[refusal])]
    client = MagicMock()
    client.responses.parse.return_value = make_client_response(None, output)
    provider = OpenAIModelProvider(make_settings(), client=client)

    with pytest.raises(LLMRefusalError, match="该请求无法完成"):
        provider.generate_deck("生成三页项目介绍", "system rules")


def test_empty_output_is_mapped_to_business_error() -> None:
    client = MagicMock()
    client.responses.parse.return_value = make_client_response(None)
    provider = OpenAIModelProvider(make_settings(), client=client)

    with pytest.raises(LLMEmptyOutputError, match="output_parsed"):
        provider.generate_deck("生成三页项目介绍", "system rules")
