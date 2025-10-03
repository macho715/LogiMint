"""종합 매퍼 로직 테스트. Comprehensive mapper logic tests."""

from __future__ import annotations

import asyncio
from pathlib import Path

import pytest

from comprehensive_email_mapper import build_summary, gather_metadata, main
from config.settings import HVDCSettings
from utils.email_parser import EmailMetadata


def test_build_summary_counts_cases() -> None:
    """요약 케이스 카운트 확인. Ensure summary case counts."""
    meta = EmailMetadata(
        path=Path("/tmp/example.eml"),
        subject="HVDC-ADOPT-HE-0476 update",
        sender="sender@example.com",
        recipients=["das@example.com"],
        cc=[],
        received_at=None,
        body="Progress for HVDC-ADOPT-HE-0476",
        attachments=[],
    )
    summary = build_summary([meta])
    assert summary["total_emails"] == 1
    assert summary["cases"][0][0] == "HVDC-ADOPT-HE-0476"


def test_gather_metadata_reads_email(tmp_path: Path) -> None:
    """비동기 메타데이터 수집 확인. Ensure metadata gathering works."""
    from tests.test_file_handler import _write_email  # reuse helper

    settings = HVDCSettings(email_root=tmp_path)
    email_path = tmp_path / "sample.eml"
    _write_email(email_path, subject="HVDC-ADOPT-HE-0476", body="Body")
    results = asyncio.run(gather_metadata(tmp_path, settings))
    assert len(results) == 1


def test_main_creates_outputs(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """메인 실행 결과 확인. Ensure main execution produces outputs."""
    from tests.test_file_handler import _write_email

    monkeypatch.setenv("HVDC_LOG_DIR", str(tmp_path))
    settings = HVDCSettings(email_root=tmp_path)
    email_path = tmp_path / "sample.eml"
    _write_email(email_path, subject="HVDC-ADOPT-HE-0476", body="Body")
    summary_path = tmp_path / "summary.json"
    jsonl_path = tmp_path / "emails.jsonl"
    exit_code = main(
        [
            "--email-root",
            str(tmp_path),
            "--output",
            str(summary_path),
            "--jsonl",
            str(jsonl_path),
        ]
    )
    assert exit_code == 0
    assert summary_path.exists()
    assert jsonl_path.exists()
