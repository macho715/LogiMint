"""화물 추적 로직 테스트. Cargo tracking logic tests."""

from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import pytest

from hvdc_cargo_tracking_system import (
    attach_email_context,
    build_cargo_records,
    determine_status,
    main,
    summarize_records,
)


def test_determine_status_site_delivery() -> None:
    """현장 인도 상태 결정 확인. Ensure site delivered status."""
    row = pd.Series({"SHU": pd.Timestamp.utcnow()})
    assert determine_status(row) == "SITE_DELIVERED"


def test_build_cargo_records_maps_status() -> None:
    """레코드 빌드 상태 매핑 확인. Ensure records map status."""
    frame = pd.DataFrame(
        [
            {
                "CASE": "HVDC-ADOPT-HE-0476",
                "ATD": pd.Timestamp.utcnow(),
                "ETA": pd.Timestamp.utcnow(),
            },
            {"CASE": "HVDC-ADOPT-SCT-0136", "SHU": pd.Timestamp.utcnow()},
        ]
    )
    records = build_cargo_records(frame)
    assert len(records) == 2
    statuses = {record.status for record in records}
    assert "운송중" in statuses or "현장인도" in statuses


def test_summarize_records_counts() -> None:
    """요약 카운트 확인. Ensure summary counts."""
    frame = pd.DataFrame(
        [
            {"CASE": "A", "SHU": pd.Timestamp.utcnow()},
            {"CASE": "B", "SHU": pd.Timestamp.utcnow()},
        ]
    )
    records = build_cargo_records(frame)
    summary = summarize_records(records)
    assert summary["total"] == 2
    assert summary["statuses"]["현장인도"] == 2


def test_attach_email_context_updates_vendor(tmp_path: Path) -> None:
    """이메일 컨텍스트 갱신 확인. Ensure email context updates vendor."""
    frame = pd.DataFrame(
        [
            {"CASE": "A", "SHU": pd.Timestamp.utcnow(), "VENDOR": None},
        ]
    )
    records = build_cargo_records(frame)
    payload = {"emails": [{"case": "A", "vendor": "Samsung C&T"}]}
    json_path = tmp_path / "emails.json"
    json_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    attach_email_context(records, json_path)
    assert records[0].vendor == "Samsung C&T"


def test_cargo_main_outputs_summary(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """메인 실행 요약 생성 확인. Ensure cargo main creates summary."""
    frame = pd.DataFrame(
        [
            {"CASE": "A", "ATD": pd.Timestamp.utcnow(), "ETA": pd.Timestamp.utcnow()},
        ]
    )
    monkeypatch.setenv("HVDC_LOG_DIR", str(tmp_path))
    status_path = tmp_path / "status.xlsx"
    status_path.write_text("dummy", encoding="utf-8")
    monkeypatch.setattr("hvdc_cargo_tracking_system.pd.read_excel", lambda path: frame.copy())
    output_path = tmp_path / "summary.json"
    exit_code = main(["--status-file", str(status_path), "--output", str(output_path)])
    assert exit_code == 0
    assert json.loads(output_path.read_text(encoding="utf-8"))["total"] == 1
