"""파일 및 이메일 유틸리티 테스트. File and email utility tests."""

from __future__ import annotations

import email
from email.message import EmailMessage
from pathlib import Path
from typing import Iterable

import pytest

from config.settings import HVDCSettings
from utils.email_parser import EmailMetadata, parse_email_file
from utils.file_handler import chunked, iter_email_files, read_bytes, safe_write_jsonl


@pytest.fixture()
def sample_settings(tmp_path: Path) -> HVDCSettings:
    """테스트 기본 설정 제공. Provide baseline test settings."""
    return HVDCSettings(
        email_root=tmp_path,
        allowed_extensions=(".eml", ".txt"),
        outlook_extensions=(".pst", ".ost", ".msg"),
        encoding_fallbacks=("utf-8", "cp1252"),
        max_files=None,
    )


def _write_email(
    path: Path, *, subject: str, body: str, attachments: Iterable[tuple[str, bytes]] | None = None
) -> None:
    """테스트용 이메일 파일 생성. Create test email file."""
    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = "sender@example.com"
    message["To"] = "receiver@example.com"
    message.set_content(body)

    for name, content in attachments or []:
        message.add_attachment(
            content, maintype="application", subtype="octet-stream", filename=name
        )

    with path.open("wb") as stream:
        stream.write(message.as_bytes())


def test_iter_email_files_excludes_outlook(tmp_path: Path, sample_settings: HVDCSettings) -> None:
    """Outlook 파일 제외 확인. Ensure Outlook files are excluded."""
    inbox = tmp_path / "Inbox"
    inbox.mkdir()
    allowed = inbox / "example.eml"
    disallowed = inbox / "archive.pst"
    allowed.write_text("dummy", encoding="utf-8")
    disallowed.write_text("binary", encoding="utf-8")

    files = list(iter_email_files(inbox, sample_settings))

    assert files == [allowed.resolve()]


def test_parse_email_file_handles_non_ascii(tmp_path: Path, sample_settings: HVDCSettings) -> None:
    """비 ASCII 이메일 파싱 보장. Guarantee parsing for non-ASCII email."""
    email_path = tmp_path / "안녕하세요.eml"
    _write_email(email_path, subject="테스트 ✉️", body="본문 내용입니다. Body text.")

    parsed = parse_email_file(email_path, sample_settings)

    assert isinstance(parsed, EmailMetadata)
    assert parsed.subject == "테스트 ✉️"
    assert "본문" in parsed.body
    assert parsed.attachments == []
    assert parsed.sender == "sender@example.com"


def test_parse_email_file_includes_attachment(
    tmp_path: Path, sample_settings: HVDCSettings
) -> None:
    """첨부 파일 메타데이터 추출. Extract attachment metadata."""
    email_path = tmp_path / "attachment.eml"
    _write_email(
        email_path, subject="첨부", body="파일 포함", attachments=[("spec.pdf", b"PDFDATA")]
    )

    parsed = parse_email_file(email_path, sample_settings)

    assert parsed.attachments[0].filename == "spec.pdf"
    assert parsed.attachments[0].size_bytes == len(b"PDFDATA")


def test_chunked_breaks_sequence() -> None:
    """청크 분할 동작 확인. Ensure chunk splitting works."""
    batches = list(chunked([1, 2, 3, 4, 5], 2))
    assert batches == [[1, 2], [3, 4], [5]]


def test_safe_write_jsonl_and_read_bytes(tmp_path: Path) -> None:
    """JSONL 기록 및 읽기 테스트. Test JSONL write and read."""
    path = tmp_path / "records.jsonl"
    safe_write_jsonl(path, [{"a": 1.2345}, {"b": "값"}])
    data = read_bytes(path)
    assert b"1.2345" in data
