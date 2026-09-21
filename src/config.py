"""应用环境配置。"""

from typing import Literal

from pydantic import PositiveFloat, PositiveInt
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
