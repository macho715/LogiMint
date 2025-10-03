"""HVDC 폴더 제목 매퍼. HVDC folder title mapper."""

from __future__ import annotations

import argparse
import json
import logging
from pathlib import Path

from config.settings import HVDCSettings, load_settings
from hvdc.logging import configure_logger
from utils.pattern_matcher import find_case_codes, match_site_alias


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    """CLI 인자 파싱. Parse CLI arguments."""
    parser = argparse.ArgumentParser(description="HVDC folder title mapper")
    parser.add_argument("--folder-root", type=Path, default=None, help="폴더 루트 경로")
    parser.add_argument("--output", type=Path, default=Path("out/folder_title_mapping.json"))
    return parser.parse_args(argv)


def build_folder_map(
    root: Path, settings: HVDCSettings, logger: logging.Logger
) -> list[dict[str, object]]:
    """폴더별 케이스 매핑. Build folder case mapping."""
    mapping: list[dict[str, object]] = []
    for folder in root.iterdir():
        if not folder.is_dir():
            continue
        codes = find_case_codes(folder.name)
        mapping.append(
            {
                "folder": folder.name,
                "path": str(folder.resolve()),
                "cases": codes,
                "site": match_site_alias(folder.name.split("_")[0]),
            }
        )
        logger.info("Mapped folder %s -> %s", folder.name, ",".join(codes))
    return mapping


def main(argv: list[str] | None = None) -> int:
    """스크립트 메인 함수. Script main function."""
    args = parse_args(argv)
    overrides = {}
    if args.folder_root is not None:
        overrides["email_root"] = args.folder_root
    settings = load_settings(overrides)
    logger = configure_logger("folder_title_mapper", settings)
    root = settings.resolve_email_root()
    if not root.exists():
        logger.error("Folder root does not exist: %s", root)
        return 1
    mapping = build_folder_map(root, settings, logger)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(mapping, ensure_ascii=False, indent=2), encoding="utf-8")
    logger.info("Saved folder mapping to %s", args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
