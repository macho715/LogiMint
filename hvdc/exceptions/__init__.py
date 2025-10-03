"""HVDC 사용자 정의 예외. HVDC custom exceptions."""

from __future__ import annotations


class HVDCError(Exception):
    """HVDC 기본 예외. HVDC base exception."""


class FileSystemError(HVDCError):
    """파일 시스템 오류. File system error."""


class EmailParsingError(HVDCError):
    """이메일 파싱 오류. Email parsing error."""


class PatternMatchError(HVDCError):
    """패턴 매칭 오류. Pattern matching error."""
