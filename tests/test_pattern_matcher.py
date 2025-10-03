"""패턴 매칭 유틸리티 테스트. Pattern matcher utility tests."""

from __future__ import annotations

from utils.pattern_matcher import (
    find_case_codes,
    match_vendor_alias,
    normalize_subject,
    score_similarity,
)


def test_normalize_subject_trims_and_lowercases() -> None:
    """주제 정규화 확인. Ensure subject normalization."""
    text = "  HVDC-ADOPT-HE-0476  \n"
    assert normalize_subject(text) == "hvdc-adopt-he-0476"


def test_find_case_codes_detects_multiple() -> None:
    """복수 케이스 매칭 확인. Ensure multiple case detection."""
    subject = "Re: HVDC-ADOPT-HE-0476 / JPTW-71 / GRM-123"
    codes = find_case_codes(subject)
    assert "HVDC-ADOPT-HE-0476" in codes
    assert "JPTW-71/GRM-123" in codes


def test_match_vendor_alias_returns_preferred() -> None:
    """벤더 별칭 매핑 확인. Verify vendor alias mapping."""
    assert match_vendor_alias("he") == "Hitachi Energy"
    assert match_vendor_alias("samsung c&t") == "Samsung C&T"


def test_score_similarity_threshold() -> None:
    """유사도 점수 범위 확인. Ensure similarity score range."""
    assert 0.0 <= score_similarity("alpha", "beta") <= 1.0
    assert score_similarity("hvdc", "hvdc") == 1.0
