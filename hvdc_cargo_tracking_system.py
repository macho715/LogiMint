"""HVDC 화물 추적 시스템. HVDC cargo tracking system."""

from __future__ import annotations

import argparse
import json
from collections import Counter
from pathlib import Path
from typing import Any, Callable, Iterable, Sequence, cast

import pandas as pd

from config.settings import HVDCSettings, load_settings
from hvdc.exceptions import HVDCError
from hvdc.logging import configure_logger
from hvdc.models import LogiBaseModel

STATUS_MAP = {
    "NOT_SHIPPED": "미출고",
    "IN_TRANSIT": "운송중",
    "PORT_ARRIVAL": "항구도착",
    "CUSTOMS": "통관중",
    "WAREHOUSE": "창고보관",
    "SITE_DELIVERED": "현장인도",
    "DELAYED": "지연",
    "UNKNOWN": "상태불명",
}

SITE_COLUMNS = ("SHU", "MIR", "DAS", "AGI")
WAREHOUSE_COLUMNS = (
    "DSV\nIndoor",
    "DSV\nOutdoor",
    "DSV\nMZD",
    "JDN\nMZD",
    "JDN\nWaterfront",
    "MOSB",
    "AAA Storage",
    "ZENER (WH)",
    "Hauler DG Storage",
    "Vijay Tanks",
)


class CargoRecord(LogiBaseModel):
    """화물 레코드 모델. Cargo record model."""

    case: str
    status: str
    vendor: str | None
    eta: str | None
    ata: str | None


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    """CLI 인자 파싱. Parse CLI arguments."""
    parser = argparse.ArgumentParser(description="HVDC cargo tracking")
    parser.add_argument("--status-file", type=Path, required=True, help="상태 엑셀 경로")
    parser.add_argument("--email-json", type=Path, default=None, help="이메일 JSON 경로")
    parser.add_argument("--output", type=Path, default=Path("out/cargo_summary.json"))
    return parser.parse_args(argv)


def determine_status(row: pd.Series) -> str:
    """상태 결정 로직. Determine cargo status."""
    for site in SITE_COLUMNS:
        if site in row and pd.notna(row[site]):
            return "SITE_DELIVERED"
    for warehouse in WAREHOUSE_COLUMNS:
        if warehouse in row and pd.notna(row[warehouse]):
            return "WAREHOUSE"
    if pd.notna(row.get("Customs\nStart")) and pd.isna(row.get("Customs\nClose")):
        return "CUSTOMS"
    if pd.notna(row.get("ATA ")):
        return "PORT_ARRIVAL"
    if pd.notna(row.get("ATD")) or pd.notna(row.get("ETD")):
        eta = row.get("ETA")
        if pd.notna(eta) and eta < pd.Timestamp.utcnow() and pd.isna(row.get("ATA ")):
            return "DELAYED"
        return "IN_TRANSIT"
    return "NOT_SHIPPED"


def load_status_dataframe(path: Path) -> pd.DataFrame:
    """상태 데이터프레임 로드. Load status dataframe."""
    if not path.exists():
        raise HVDCError(f"Status file not found: {path}")
    frame = pd.read_excel(path)
    date_columns = ["ETD", "ATD", "ETA", "ATA "] + list(SITE_COLUMNS) + list(WAREHOUSE_COLUMNS)
    for column in date_columns:
        if column in frame.columns:
            frame[column] = pd.to_datetime(frame[column], errors="coerce")
    return frame


def build_cargo_records(frame: pd.DataFrame) -> list[CargoRecord]:
    """데이터프레임에서 화물 레코드 생성. Build cargo records from dataframe."""
    records: list[CargoRecord] = []
    for _, row in frame.iterrows():
        status_code = determine_status(row)
        records.append(
            CargoRecord(
                case=str(row.get("CASE", "")),
                status=STATUS_MAP.get(status_code, STATUS_MAP["UNKNOWN"]),
                vendor=row.get("VENDOR"),
                eta=_to_iso(row.get("ETA")),
                ata=_to_iso(row.get("ATA ")),
            )
        )
    return records


def _to_iso(value: object) -> str | None:
    """ISO 문자열 변환. Convert to ISO string."""
    if isinstance(value, pd.Timestamp):
        return cast(str, value.to_pydatetime().isoformat())
    iso_method = getattr(value, "isoformat", None)
    if callable(iso_method):
        try:
            method = cast(Callable[[], Any], iso_method)
            result = method()
            if isinstance(result, str):
                return result
            text = str(result)
            return text
        except Exception:  # pragma: no cover - defensive
            return None
    return None


def summarize_records(records: Iterable[CargoRecord]) -> dict[str, object]:
    """요약 정보 생성. Build summary information."""
    counter = Counter(record.status for record in records)
    return {
        "total": sum(counter.values()),
        "statuses": counter,
    }


def attach_email_context(records: list[CargoRecord], email_path: Path | None) -> None:
    """이메일 컨텍스트 연결. Attach email context."""
    if email_path is None or not email_path.exists():
        return
    with email_path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)
    lookup = {entry.get("case"): entry for entry in payload.get("emails", [])}
    for record in records:
        context = lookup.get(record.case)
        if context:
            record.vendor = context.get("vendor", record.vendor)


def main(argv: Sequence[str] | None = None) -> int:
    """메인 실행 함수. Main execution function."""
    args = parse_args(argv)
    settings = load_settings({})
    logger = configure_logger("hvdc_cargo_tracking_system", settings)
    try:
        frame = load_status_dataframe(args.status_file)
        records = build_cargo_records(frame)
        attach_email_context(records, args.email_json)
    except HVDCError as error:
        logger.error("Cargo tracking failed: %s", error)
        return 1
    summary = summarize_records(records)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    logger.info("Cargo summary stored in %s", args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
