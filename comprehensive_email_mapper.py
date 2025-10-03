"""HVDC 종합 이메일 매퍼. HVDC comprehensive email mapper."""

from __future__ import annotations

import argparse
import asyncio
import json
from collections import Counter
from pathlib import Path
from typing import Iterable, Sequence

from config.settings import HVDCSettings, load_settings
from hvdc.exceptions import EmailParsingError, HVDCError
from hvdc.logging import configure_logger
from utils.email_parser import EmailMetadata, parse_email_file
from utils.file_handler import chunked, iter_email_files, safe_write_jsonl
from utils.pattern_matcher import find_case_codes, match_site_alias, match_vendor_alias


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    """CLI 인자 파싱. Parse CLI arguments."""
    parser = argparse.ArgumentParser(description="HVDC comprehensive email mapper")
    parser.add_argument("--email-root", type=Path, default=None, help="이메일 루트 경로")
    parser.add_argument(
        "--output", type=Path, default=Path("out/comprehensive_summary.json"), help="요약 출력 파일"
    )
    parser.add_argument(
        "--jsonl",
        type=Path,
        default=Path("out/comprehensive_emails.jsonl"),
        help="이메일 JSONL 경로",
    )
    return parser.parse_args(argv)


async def _parse_email(path: Path, settings: HVDCSettings) -> EmailMetadata | None:
    """이메일 비동기 파싱. Parse email asynchronously."""
    try:
        return await asyncio.to_thread(parse_email_file, path, settings)
    except EmailParsingError:
        return None


async def gather_metadata(root: Path, settings: HVDCSettings) -> list[EmailMetadata]:
    """메타데이터 수집. Gather metadata."""
    sem = asyncio.Semaphore(settings.async_batch_size)
    results: list[EmailMetadata] = []

    async def worker(target: Path) -> None:
        async with sem:
            record = await _parse_email(target, settings)
            if record is not None:
                results.append(record)

    tasks: list[asyncio.Task[None]] = []
    for batch in chunked(iter_email_files(root, settings), settings.async_batch_size):
        tasks.extend(asyncio.create_task(worker(item)) for item in batch)
    if tasks:
        await asyncio.gather(*tasks)
    return results


def build_summary(items: Iterable[EmailMetadata]) -> dict[str, object]:
    """요약 데이터 생성. Build summary data."""
    vendor_counter: Counter[str] = Counter()
    site_counter: Counter[str] = Counter()
    case_counter: Counter[str] = Counter()

    total = 0
    for item in items:
        total += 1
        codes = find_case_codes(f"{item.subject} {item.body[:200]}")
        for code in codes:
            case_counter[code] += 1
            vendor_parts = code.split("-")
            if vendor_parts:
                vendor_counter[match_vendor_alias(vendor_parts[1])] += 1
        for recipient in item.recipients:
            site_counter[match_site_alias(recipient.split("@")[0])] += 1

    return {
        "total_emails": total,
        "cases": case_counter.most_common(),
        "vendors": vendor_counter.most_common(),
        "sites": site_counter.most_common(),
    }


async def main_async(args: argparse.Namespace) -> int:
    """비동기 메인 실행. Run asynchronous main."""
    overrides = {}
    if args.email_root is not None:
        overrides["email_root"] = args.email_root
    settings = load_settings(overrides)
    logger = configure_logger("comprehensive_email_mapper", settings)
    root = settings.resolve_email_root()
    logger.info("Scanning emails from %s", root)
    metadata = await gather_metadata(root, settings)
    logger.info("Parsed %d emails", len(metadata))
    if not metadata:
        logger.warning("No emails parsed")
    summary = build_summary(metadata)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    safe_write_jsonl(args.jsonl, (item.model_dump_safe() for item in metadata))
    logger.info("Summary saved to %s", args.output)
    return 0


def main(argv: Sequence[str] | None = None) -> int:
    """동기 메인 래퍼. Synchronous main wrapper."""
    args = parse_args(argv)
    try:
        return asyncio.run(main_async(args))
    except HVDCError as error:
        print(f"Error: {error}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
