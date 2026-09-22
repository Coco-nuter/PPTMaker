"""模型提供商适配器与领域错误。"""

from __future__ import annotations

from typing import Any, Protocol

from openai import APIConnectionError, APITimeoutError, OpenAI, OpenAIError
from pydantic import BaseModel, ValidationError

from config import Settings
from models import DeckSpec


class LLMError(RuntimeError):
    """模型规划失败的业务异常基类。"""


class LLMConfigurationError(LLMError):
    """模型调用配置缺失或无效。"""


class MissingAPIKeyError(LLMConfigurationError):
    """没有配置 OpenAI API Key。"""


class LLMTimeoutError(LLMError):
    """模型请求超过配置的时间限制。"""


class LLMNetworkError(LLMError):
    """模型请求因网络连接失败。"""


class LLMRefusalError(LLMError):
    """模型明确拒绝完成请求。"""


class LLMEmptyOutputError(LLMError):
    """模型没有返回可解析的结构化结果。"""


class LLMInvalidOutputError(LLMError):
    """模型结果未通过 DeckSpec 校验。"""


class LLMProviderError(LLMError):
    """未被单独分类的模型提供商错误。"""


class ModelProvider(Protocol):
    """规划器依赖的最小领域接口，不暴露 OpenAI SDK 类型。"""

    def generate_deck(self, user_request: str, system_prompt: str) -> DeckSpec:
        """把自然语言要求转换为合法 DeckSpec。"""
        ...


class FakeModelProvider:
    """供自动测试使用的确定性 DeckSpec provider。"""

    def __init__(
        self,
        decks: DeckSpec | list[DeckSpec],
        *,
        error: Exception | None = None,
    ) -> None:
        self._decks = list(decks) if isinstance(decks, list) else [decks]
        if not self._decks:
            raise ValueError("FakeModelProvider requires at least one DeckSpec")
        self.error = error
        self.calls: list[tuple[str, str]] = []
        self._next_index = 0

    def generate_deck(self, user_request: str, system_prompt: str) -> DeckSpec:
        """记录调用，按顺序返回合法副本，耗尽后重复最后一个结果。"""
        self.calls.append((user_request, system_prompt))
        if self.error is not None:
            raise self.error

        index = min(self._next_index, len(self._decks) - 1)
        self._next_index += 1
        return validate_deck_output(self._decks[index])


def _read_value(value: object, name: str) -> object | None:
    """同时读取 SDK 对象或测试字典中的字段。"""
    if isinstance(value, dict):
        return value.get(name)
    return getattr(value, name, None)


def _find_refusal(response: object) -> str | None:
    """从 Responses API 输出项中提取明确拒绝原因。"""
    output = _read_value(response, "output")
    if not isinstance(output, (list, tuple)):
        return None

    for item in output:
        content = _read_value(item, "content")
        if not isinstance(content, (list, tuple)):
            continue
        for part in content:
            if _read_value(part, "type") != "refusal":
                continue
            reason = _read_value(part, "refusal")
            if isinstance(reason, str) and reason.strip():
                return reason.strip()
            return "模型拒绝了该请求，但未提供具体原因。"
    return None


def validate_deck_output(parsed: object) -> DeckSpec:
    """把 SDK 解析结果转换成普通数据后再次执行完整领域校验。"""
    raw_output: object
    if isinstance(parsed, BaseModel):
        raw_output = parsed.model_dump(mode="json")
    else:
        raw_output = parsed

    try:
        return DeckSpec.model_validate(raw_output)
    except (TypeError, ValidationError) as exc:
        raise LLMInvalidOutputError(f"模型返回的 DeckSpec 不合法：{exc}") from exc


class OpenAIModelProvider:
    """使用 OpenAI Responses API 结构化输出生成 DeckSpec。"""

    def __init__(self, settings: Settings | None = None, client: Any | None = None) -> None:
        self._settings = settings or Settings()
        api_key = self._require_api_key(self._settings)
        self._model = self._require_model(self._settings.openai_model)

        if client is None:
            client_options: dict[str, object] = {
                "api_key": api_key,
                "timeout": float(self._settings.llm_timeout_seconds),
                "max_retries": 1,
            }
            if self._settings.openai_base_url is not None:
                client_options["base_url"] = self._settings.openai_base_url
            client = OpenAI(**client_options)
        self._client = client

    @staticmethod
    def _require_api_key(settings: Settings) -> str:
        api_key = settings.openai_api_key
        if api_key is None or not api_key.get_secret_value().strip():
            raise MissingAPIKeyError(
                "缺少 OPENAI_API_KEY；请在本地 .env 或当前 PowerShell 会话中配置。"
            )
        return api_key.get_secret_value()

    @staticmethod
    def _require_model(model: str | None) -> str:
        if model is None or not model.strip():
            raise LLMConfigurationError("缺少 OPENAI_MODEL；请配置支持结构化输出的模型。")
        return model.strip()

    def generate_deck(self, user_request: str, system_prompt: str) -> DeckSpec:
        """调用结构化输出接口，并将 SDK 错误映射为可操作的业务错误。"""
        try:
            response = self._client.responses.parse(
                model=self._model,
                input=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_request},
                ],
                text_format=DeckSpec,
                store=False,
                timeout=float(self._settings.llm_timeout_seconds),
            )
        except APITimeoutError as exc:
            raise LLMTimeoutError(
                f"OpenAI 请求超过 {self._settings.llm_timeout_seconds:g} 秒。"
            ) from exc
        except APIConnectionError as exc:
            raise LLMNetworkError("无法连接 OpenAI API，请检查网络和 OPENAI_BASE_URL。") from exc
        except ValidationError as exc:
            raise LLMInvalidOutputError(f"模型返回内容无法解析为 DeckSpec：{exc}") from exc
        except OpenAIError as exc:
            raise LLMProviderError(f"OpenAI API 调用失败：{exc.__class__.__name__}") from exc

        parsed = _read_value(response, "output_parsed")
        if parsed is None:
            refusal = _find_refusal(response)
            if refusal is not None:
                raise LLMRefusalError(f"模型拒绝生成演示文稿：{refusal}")
            raise LLMEmptyOutputError("OpenAI 未返回 output_parsed，无法生成 DeckSpec。")

        return validate_deck_output(parsed)
