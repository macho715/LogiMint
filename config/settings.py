"""HVDC 전역 설정 모듈. HVDC global settings module."""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, Iterable

from pydantic import Field

from hvdc.models import LogiBaseModel

DEFAULT_ALLOWED_EXT = (".eml", ".txt", ".html", ".csv")
DEFAULT_OUTLOOK_EXT = (".pst", ".ost", ".msg")
DEFAULT_ENCODINGS = ("utf-8", "cp1252", "latin-1")


def _env_bool(name: str, default: str) -> bool:
    """환경 불리언 파싱. Parse environment boolean."""
    return bool(int(os.getenv(name, default)))


def _env_int(name: str, default: str) -> int:
    """환경 정수 파싱. Parse environment integer."""
    return int(os.getenv(name, default))


class HVDCSettings(LogiBaseModel):
    """HVDC 설정 데이터 클래스. HVDC settings data class."""

    email_root: Path = Field(default_factory=lambda: Path(os.getenv("HVDC_EMAIL_ROOT", "./emails")))
    allowed_extensions: tuple[str, ...] = Field(default=DEFAULT_ALLOWED_EXT)
    outlook_extensions: tuple[str, ...] = Field(default=DEFAULT_OUTLOOK_EXT)
    encoding_fallbacks: tuple[str, ...] = Field(default=DEFAULT_ENCODINGS)
    max_files: int | None = Field(default=None)
    log_directory: Path = Field(default_factory=lambda: Path(os.getenv("HVDC_LOG_DIR", "./logs")))
    log_json: bool = Field(default_factory=lambda: _env_bool("HVDC_LOG_JSON", "0"))
    log_rotation_mb: int = Field(default_factory=lambda: _env_int("HVDC_LOG_ROTATION_MB", "5"))
    log_retention: int = Field(default_factory=lambda: _env_int("HVDC_LOG_RETENTION", "5"))
    async_batch_size: int = Field(default_factory=lambda: _env_int("HVDC_ASYNC_BATCH_SIZE", "32"))
    excel_chunk_size: int = Field(default_factory=lambda: _env_int("HVDC_EXCEL_CHUNK_SIZE", "500"))

    def resolve_email_root(self) -> Path:
        """이메일 루트 경로 반환. Return email root path."""
        return self.email_root.expanduser().resolve()

    def to_json(self) -> str:
        """JSON 문자열 반환. Return JSON string."""
        return json.dumps(self.model_dump(mode="json"), ensure_ascii=False)


def load_settings(overrides: dict[str, Any] | None = None) -> HVDCSettings:
    """환경 기반 설정 로드. Load settings from environment."""
    base = HVDCSettings()
    if not overrides:
        return base
    return base.model_copy(update=overrides)


def ensure_directory(path: Path) -> Path:
    """디렉터리 생성 보장. Ensure directory exists."""
    path.mkdir(parents=True, exist_ok=True)
    return path


def merge_extensions(primary: Iterable[str], extra: Iterable[str]) -> tuple[str, ...]:
    """확장자 병합 헬퍼. Merge file extensions helper."""
    normalized = {ext.lower() if ext.startswith(".") else f".{ext.lower()}" for ext in primary}
    for ext in extra:
        normalized.add(ext.lower() if ext.startswith(".") else f".{ext.lower()}")
    return tuple(sorted(normalized))
