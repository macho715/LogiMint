"""이메일 파서 유틸리티. Email parser utilities."""

from __future__ import annotations

import email
from email import policy
from email.header import decode_header
from email.message import EmailMessage, Message
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Iterable, List, cast

from config.settings import HVDCSettings
from hvdc.exceptions import EmailParsingError
from hvdc.models import LogiBaseModel
from utils.file_handler import read_bytes


class AttachmentMetadata(LogiBaseModel):
    """첨부 메타데이터. Attachment metadata."""

    filename: str
    content_type: str
    size_bytes: int


class EmailMetadata(LogiBaseModel):
    """이메일 메타데이터. Email metadata."""

    path: Path
    subject: str
    sender: str
    recipients: list[str]
    cc: list[str]
    received_at: str | None
    body: str
    attachments: list[AttachmentMetadata]


def _decode_header(raw_header: str | None, encodings: Iterable[str]) -> str:
    """헤더 디코드 헬퍼. Helper to decode headers."""
    if not raw_header:
        return ""
    decoded_parts: List[str] = []
    for fragment, encoding in decode_header(raw_header):
        if isinstance(fragment, str):
            decoded_parts.append(fragment)
            continue
        candidates = list(encodings) + [encoding or "utf-8"]
        for candidate in candidates:
            try:
                decoded_parts.append(fragment.decode(candidate or "utf-8"))
                break
            except (LookupError, UnicodeDecodeError):
                continue
        else:
            decoded_parts.append(fragment.decode("utf-8", errors="ignore"))
    return "".join(decoded_parts).strip()


def _extract_body(message: Message, encodings: Iterable[str]) -> str:
    """본문 텍스트 추출. Extract body text."""
    if message.is_multipart():
        for part in message.walk():
            if part.is_multipart():
                continue
            content_type = part.get_content_type()
            if content_type not in {"text/plain", "text/html"}:
                continue
            payload = _decode_payload(part, encodings)
            if payload:
                return payload
        return ""
    return _decode_payload(message, encodings)


def _decode_payload(part: Message, encodings: Iterable[str]) -> str:
    """페이로드 디코드. Decode payload."""
    payload = cast(bytes | None, part.get_payload(decode=True))
    if payload is None:
        text = part.get_payload()
        return text if isinstance(text, str) else ""
    for encoding in encodings:
        try:
            return payload.decode(encoding)
        except (LookupError, UnicodeDecodeError):
            continue
    return payload.decode("utf-8", errors="ignore")


def _extract_attachments(message: Message) -> list[AttachmentMetadata]:
    """첨부 리스트 추출. Extract attachments list."""
    attachments: list[AttachmentMetadata] = []
    iter_attachments = getattr(message, "iter_attachments", None)
    if callable(iter_attachments):
        for attachment in iter_attachments():
            payload = attachment.get_payload(decode=True) or b""
            attachments.append(
                AttachmentMetadata(
                    filename=attachment.get_filename() or "attachment.bin",
                    content_type=attachment.get_content_type(),
                    size_bytes=len(payload),
                )
            )
    return attachments


def parse_email_file(path: Path, settings: HVDCSettings) -> EmailMetadata:
    """이메일 파일을 파싱. Parse email file."""
    try:
        raw_bytes = read_bytes(path)
        message: EmailMessage = email.message_from_bytes(raw_bytes, policy=policy.default)
        encodings = settings.encoding_fallbacks
        subject = _decode_header(message.get("Subject"), encodings)
        sender = _decode_header(message.get("From"), encodings)
        to_list = _decode_header(message.get("To"), encodings)
        cc_list = _decode_header(message.get("Cc"), encodings)
        received = message.get("Date")
        received_at = None
        if received:
            try:
                received_at = parsedate_to_datetime(received).isoformat()
            except (TypeError, ValueError):
                received_at = None
        body = _extract_body(message, encodings)
        attachments = _extract_attachments(message)
        recipients = [addr.strip() for addr in to_list.split(",") if addr.strip()]
        cc = [addr.strip() for addr in cc_list.split(",") if addr.strip()]
        return EmailMetadata(
            path=path.resolve(),
            subject=subject,
            sender=sender,
            recipients=recipients,
            cc=cc,
            received_at=received_at,
            body=body,
            attachments=attachments,
        )
    except EmailParsingError:
        raise
    except Exception as error:  # pylint: disable=broad-except
        raise EmailParsingError(str(error)) from error
