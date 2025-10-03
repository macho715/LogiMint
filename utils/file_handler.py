"""파일 시스템 유틸리티. File system utilities."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Iterable, Iterator, Sequence, TypeVar

from config.settings import HVDCSettings
from hvdc.exceptions import FileSystemError

T = TypeVar("T")


def iter_email_files(root: Path, settings: HVDCSettings) -> Iterator[Path]:
    """이메일 파일 생성기. Yield email files generator."""
    try:
        allowed = {ext.lower() for ext in settings.allowed_extensions}
        blocked = {ext.lower() for ext in settings.outlook_extensions}
        yielded = 0
        for candidate in root.rglob("*"):
            if not candidate.is_file():
                continue
            suffix = candidate.suffix.lower()
            if suffix in blocked:
                continue
            if suffix not in allowed:
                continue
            yield candidate.resolve()
            yielded += 1
            if settings.max_files is not None and yielded >= settings.max_files:
                break
    except OSError as error:
        raise FileSystemError(str(error)) from error


def chunked(iterable: Sequence[T] | Iterable[T], size: int) -> Iterator[list[T]]:
    """시퀀스를 청크로 분할. Split iterable into chunks."""
    batch: list[T] = []
    for item in iterable:
        batch.append(item)
        if len(batch) >= size:
            yield batch
            batch = []
    if batch:
        yield batch


def safe_write_jsonl(path: Path, records: Iterable[dict[str, object]]) -> None:
    """JSONL 안전 기록. Safely write JSONL records."""
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", encoding="utf-8") as stream:
            for record in records:
                stream.write(json.dumps(record, ensure_ascii=False) + "\n")
    except OSError as error:
        raise FileSystemError(str(error)) from error


def read_bytes(path: Path) -> bytes:
    """파일 바이트 읽기. Read file bytes."""
    try:
        return path.read_bytes()
    except OSError as error:
        raise FileSystemError(str(error)) from error
