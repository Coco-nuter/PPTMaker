"""应用环境配置。"""

from typing import Literal

from pydantic import PositiveFloat, PositiveInt, SecretStr, field_validator
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """从进程环境和本地 .env 读取应用配置。"""

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=False,
        extra="ignore",
    )

    preview_backend: Literal["powerpoint"] = "powerpoint"
    preview_timeout_seconds: PositiveFloat = 120
    preview_width: PositiveInt = 1920
    preview_height: PositiveInt = 1080

    openai_api_key: SecretStr | None = None
    openai_base_url: str | None = None
    openai_model: str | None = None
    llm_timeout_seconds: PositiveFloat = 60

    @field_validator("openai_base_url", "openai_model", mode="before")
    @classmethod
    def blank_optional_text_is_none(cls, value: object) -> object:
        """将 `.env` 中的空占位符视为未配置。"""
        if isinstance(value, str):
            stripped = value.strip()
            return stripped or None
        return value

    @field_validator("openai_api_key", mode="before")
    @classmethod
    def blank_api_key_is_none(cls, value: object) -> object:
        """避免把空字符串误当成已配置的密钥。"""
        if isinstance(value, str) and not value.strip():
            return None
        return value
