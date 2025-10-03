"""패턴 정의 모듈. Pattern definition module."""

from __future__ import annotations

import re
from typing import Mapping

PATTERN_DEFINITIONS: Mapping[str, str] = {
    "HVDC_ADOPT": r"HVDC-ADOPT-[A-Z]+-[0-9]{3,5}",
    "HVDC_GENERIC": r"HVDC-[A-Z]+-[A-Z0-9]+-[A-Z0-9\-]+",
    "JPTW_GRM": r"JPTW-(\d+)\s*/\s*GRM-(\d+)",
    "PAREN_ANY": r"\(([^\)]+)\)",
    "TRAILING": r":\s*([A-Z]+(?:-[A-Z0-9]+){2,})",
}

COMPILED_PATTERNS = {
    key: re.compile(pattern, re.IGNORECASE) for key, pattern in PATTERN_DEFINITIONS.items()
}

VENDOR_ALIASES: Mapping[str, str] = {
    "he": "Hitachi Energy",
    "hitachi energy": "Hitachi Energy",
    "hitachi": "Hitachi Energy",
    "sct": "Samsung C&T",
    "samsung c&t": "Samsung C&T",
    "mosb": "MOSB",
    "als": "ALS",
}

SITE_ALIASES: Mapping[str, str] = {
    "das": "DAS Site",
    "agi": "AGI Site",
    "mir": "MIR Site",
    "mirfa": "MIRFA Site",
    "ghallan": "GHALLAN Site",
}
