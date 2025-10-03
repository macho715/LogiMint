"""폴더 매퍼 테스트. Folder mapper tests."""

from __future__ import annotations

import logging
from pathlib import Path

import pytest

from config.settings import HVDCSettings
from folder_title_mapper import build_folder_map, main


def test_build_folder_map_extracts_codes(tmp_path: Path) -> None:
    """폴더 코드 추출 검증. Ensure folder code extraction."""
    folder = tmp_path / "HVDC-ADOPT-HE-0476_docs"
    folder.mkdir()
    settings = HVDCSettings(email_root=tmp_path)
    logger = logging.getLogger("test_folder_mapper")
    mapping = build_folder_map(tmp_path, settings, logger)
    assert mapping[0]["cases"][0] == "HVDC-ADOPT-HE-0476"


def test_main_generates_mapping(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """메인 실행 매핑 검증. Ensure main execution produces mapping."""
    folder = tmp_path / "HVDC-ADOPT-HE-0476_docs"
    folder.mkdir()
    monkeypatch.setenv("HVDC_EMAIL_ROOT", str(tmp_path))
    monkeypatch.setenv("HVDC_LOG_DIR", str(tmp_path))
    output = tmp_path / "mapping.json"
    exit_code = main(["--folder-root", str(tmp_path), "--output", str(output)])
    assert exit_code == 0
    assert output.exists()
