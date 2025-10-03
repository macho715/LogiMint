"""패턴 매칭 유틸리티. Pattern matching utilities."""

from __future__ import annotations

import difflib
import unicodedata
from typing import Iterable

from config.patterns import COMPILED_PATTERNS, SITE_ALIASES, VENDOR_ALIASES
from hvdc.exceptions import PatternMatchError


def normalize_subject(subject: str) -> str:
    """제목 문자열 정규화. Normalize subject string."""
    normalized = unicodedata.normalize("NFKC", subject or "")
    return normalized.strip().lower()


def find_case_codes(text: str) -> list[str]:
    """케이스 코드 찾기. Find case codes."""
    try:
        work = normalize_subject(text)
        matches: set[str] = set()
        adopt = [value.upper() for value in COMPILED_PATTERNS["HVDC_ADOPT"].findall(work)]
        matches.update(adopt)
        generic = [value.upper() for value in COMPILED_PATTERNS["HVDC_GENERIC"].findall(work)]
        matches.update(generic)
        for match in COMPILED_PATTERNS["JPTW_GRM"].finditer(work):
            matches.add(f"JPTW-{match.group(1)}/GRM-{match.group(2)}")
        for match in COMPILED_PATTERNS["TRAILING"].findall(work):
            matches.add(match.upper())
        return sorted(matches)
    except KeyError as error:
        raise PatternMatchError(str(error)) from error


def match_vendor_alias(alias: str) -> str:
    """벤더 별칭 정규화. Normalize vendor alias."""
    key = normalize_subject(alias)
    return VENDOR_ALIASES.get(key, alias.strip().title())


def match_site_alias(alias: str) -> str:
    """사이트 별칭 정규화. Normalize site alias."""
    key = normalize_subject(alias)
    return SITE_ALIASES.get(key, alias.strip().upper())


def score_similarity(left: str, right: str) -> float:
    """유사도 점수 계산. Compute similarity score."""
    return round(
        difflib.SequenceMatcher(a=normalize_subject(left), b=normalize_subject(right)).ratio(), 4
    )


def fuzzy_contains(term: str, candidates: Iterable[str], threshold: float = 0.82) -> bool:
    """퍼지 포함 확인. Check fuzzy containment."""
    for candidate in candidates:
        if score_similarity(term, candidate) >= threshold:
            return True
    return False
