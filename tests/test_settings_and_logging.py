"""설정 및 로깅 테스트. Settings and logging tests."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from config.settings import HVDCSettings, ensure_directory, load_settings, merge_extensions
from hvdc.logging import configure_logger


def test_load_settings_override(tmp_path: Path) -> None:
    """오버라이드 동작 확인. Ensure overrides apply."""
    settings = load_settings({"email_root": tmp_path})
    assert settings.email_root == tmp_path


def test_configure_logger_creates_files(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """로거 생성 시 파일 생성 확인. Logger should create file."""
    monkeypatch.setenv("HVDC_LOG_DIR", str(tmp_path))
    monkeypatch.setenv("HVDC_LOG_JSON", "1")
    settings = HVDCSettings()
    logger = configure_logger("test", settings)
    logger.info("hello")
    log_file = tmp_path / "test.log"
    assert log_file.exists()
    content = log_file.read_text(encoding="utf-8").strip()
    data = json.loads(content)
    assert data["message"] == "hello"


def test_merge_extensions(tmp_path: Path) -> None:
    """확장자 병합 기능 확인. Ensure extension merging."""
    merged = merge_extensions([".eml"], ["MSG", ".txt"])
    assert merged == (".eml", ".msg", ".txt")
    created = ensure_directory(tmp_path / "logs")
    assert created.exists()
