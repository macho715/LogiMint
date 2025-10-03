# S1) core 신설 (행위 불변, 신규 파일만 추가)

## 1) 디렉터리

```
hvdc/
  core/
    __init__.py
    config.py
    errors.py
    regex.py
    io.py
    typing.py
```

## 2) 내용

### `hvdc/core/config.py`

```python
from __future__ import annotations
from pathlib import Path

# 중앙 설정(기존 하드코딩 흡수 예정) — 지금은 기본값만
EMAIL_ROOT = Path("EMAIL")                 # 예: C:\Users\...\Documents\EMAIL
ALLOWED_EXT = {".eml", ".txt", ".html", ".csv"}  # 스캐너 화이트리스트
EXCEL_OUTDIR = Path("out")
EXCEL_FILENAME = "hvdc_email_report.xlsx"

# 실행 파라미터(샘플 제한 등)
MAX_FILES = None  # None = 제한 없음
ENCODING_FALLBACKS = ["utf-8", "cp1252", "latin-1"]

# 로깅(단순 프리셋; 실제 로거는 각 엔트리에서 구성)
LOG_LEVEL = "INFO"
```

### `hvdc/core/errors.py`

```python
class HvdcError(Exception): ...
class ScanError(HvdcError): ...
class ParseError(HvdcError): ...
class PatternError(HvdcError): ...
class IoError(HvdcError): ...
```

### `hvdc/core/regex.py`

```python
import re

# 단일 카탈로그 — 기존 스크립트에서 쓰던 변형을 포괄
# 필요시 tests로 스냅샷 고정
PATTERNS = {
    # 예) HVDC-ADOPT-HE-0476, HVDC-ADOPT-SCT-0136 등
    "HVDC_ADOPT": r"HVDC-ADOPT-[A-Z]+-[0-9]{3,5}",
    # 보다 일반화된 HVDC 코드(세그먼트 3~4개)
    "HVDC_GENERIC": r"HVDC-[A-Z]+-[A-Z]+-[0-9A-Z\-]+",
    # PRL 코드(예시): PRL-O-046-O4(HE-0486)류
    "PRL": r"PRL-[A-Z]-\d{2,4}-[A-Z0-9]+(?:\([A-Z]{2}-\d{3,5}\))?",
    # JPTW/GRM 페어 추출: ... JPTW-71 / GRM-123 ...
    "JPTW_GRM": r"JPTW-(\d+)\s*/\s*GRM-(\d+)",
    # 괄호 벤더/부가 정보
    "PAREN_ANY": r"\(([^\)]+)\)",
    # 뒤따르는 트레일러 식별자(완충)
    "TRAILING": r":\s*([A-Z]+(?:-[A-Z0-9]+){2,})",
}

# 사전 컴파일(성능·일관)
COMPILED = {k: re.compile(v, re.IGNORECASE) for k, v in PATTERNS.items()}
```

### `hvdc/core/io.py`

```python
from __future__ import annotations
from pathlib import Path
from .errors import IoError

def read_text(path: Path, encodings: list[str]) -> str:
    last_err = None
    for enc in encodings:
        try:
            return path.read_text(encoding=enc, errors="strict")
        except Exception as e:
            last_err = e
    # 최종 실패 시
    raise IoError(f"failed to read {path}") from last_err

def ensure_dir(path: Path) -> None:
    try:
        path.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        raise IoError(f"failed to create dir {path}") from e
```

### `hvdc/core/typing.py`

```python
from __future__ import annotations
from typing import TypedDict, NotRequired, List

class CaseHit(TypedDict):
    value: str
    kind: str        # "HVDC_ADOPT" | "HVDC_GENERIC" | "PRL" | "JPTW" | "GRM"
    span: NotRequired[tuple[int, int]]

class ParseResult(TypedDict):
    cases: List[CaseHit]
    sites: List[str]
    lpos: List[str]
    phases: List[str]
```

---

# S2) extractors 통합 (행위 불변, 최소 기능)

## 1) 디렉터리

```
hvdc/
  extractors/
    __init__.py
    case.py
    site.py
    lpo.py
    phase.py
```

## 2) 내용

### `hvdc/extractors/case.py`

```python
from __future__ import annotations
from typing import List
from ..core.regex import COMPILED
from ..core.typing import CaseHit

def extract_cases(text: str) -> List[CaseHit]:
    """제목/본문/폴더명에서 케이스류 식별자를 추출(중복 제거, 안정 정렬)."""
    hits: list[CaseHit] = []

    def add(kind: str, value: str, span: tuple[int, int] | None):
        hits.append({"value": value, "kind": kind, "span": span})

    # HVDC ADOPT
    for m in COMPILED["HVDC_ADOPT"].finditer(text):
        add("HVDC_ADOPT", m.group(0), m.span())

    # 일반 HVDC
    for m in COMPILED["HVDC_GENERIC"].finditer(text):
        val = m.group(0)
        # ADOPT와 중복되는 경우는 스킵
        if not val.upper().startswith("HVDC-ADOPT-"):
            add("HVDC_GENERIC", val, m.span())

    # PRL
    for m in COMPILED["PRL"].finditer(text):
        add("PRL", m.group(0), m.span())

    # JPTW / GRM (페어)
    for m in COMPILED["JPTW_GRM"].finditer(text):
        add("JPTW", f"JPTW-{m.group(1)}", m.span(1))
        add("GRM",  f"GRM-{m.group(2)}", m.span(2))

    # 중복 제거(값+kind 기준) & 입력 순서 안정
    seen = set()
    dedup: list[CaseHit] = []
    for h in hits:
        key = (h["kind"].upper(), h["value"].upper())
        if key in seen:
            continue
        seen.add(key)
        dedup.append(h)

    return dedup
```

### `hvdc/extractors/site.py`

```python
from __future__ import annotations
from typing import List
import re

# 간단한 자리표시자 — 이후 rules.json/테이블로 이관 가능
SITE_TOKENS = ["AGI", "DAS", "MIRFA", "SHU", "ZAK"]

def extract_sites(text: str) -> List[str]:
    found = []
    upper = text.upper()
    for tok in SITE_TOKENS:
        if re.search(rf"\b{tok}\b", upper):
            found.append(tok)
    return sorted(set(found), key=found.index)
```

### `hvdc/extractors/lpo.py`

```python
from __future__ import annotations
from typing import List
import re

# 예: LPO-12345, PO NO : 5001005009 등도 확장 가능
LPO_RE = re.compile(r"\b(?:LPO|PO)\s*[-:]?\s*(\d{5,12})\b", re.IGNORECASE)

def extract_lpos(text: str) -> List[str]:
    vals = [m.group(1) for m in LPO_RE.finditer(text)]
    # 'PO 5001005009' → 'PO-5001005009'로 정규화
    norm = [f"PO-{v}" for v in vals]
    seen = set()
    out: list[str] = []
    for v in norm:
        if v not in seen:
            seen.add(v)
            out.append(v)
    return out
```

### `hvdc/extractors/phase.py`

```python
from __future__ import annotations
from typing import List
import re

# 예시: PHASE-1, PH1, STAGE-2 등(실제 규칙은 이후 확장)
PHASE_RE = re.compile(r"\b(?:PHASE|PH|STAGE)[- ]?(\d{1,2})\b", re.IGNORECASE)

def extract_phases(text: str) -> List[str]:
    return [f"PHASE-{m.group(1)}" for m in PHASE_RE.finditer(text)]
```

---

## 커밋 메시지 (Conventional + structural 스코프)

> 각 커밋은 **작동 불변**(기존 엔트리포인트 비변경)이고, 새 모듈만 추가합니다.

1. `structural(core): scaffold core modules (config, errors, regex, io, typing)`

   * 목적: 공통 설정·예외·정규식 카탈로그·I/O 유틸·타입 정의 신설
   * 노트: 정규식은 카탈로그·사전컴파일 구조로 통합(기능 변동 없음, 아직 미사용)

2. `structural(extractors): add unified case/site/lpo/phase extractors (pure functions)`

   * 목적: 분산된 패턴/추출 로직의 단일 진입점 마련
   * 노트: `extract_cases()`는 ADOPT/GENERIC/PRL/JPTW·GRM 중복 제거 규칙만 포함(출력 안정성 우선)

> 선택(테스트 동반 커밋 원하면):

3. `tests(unit): add initial snapshot tests for core.regex and extractors.case`

   * 예: 대표 제목 10–20개로 PATTERN 매칭 스냅샷 고정

---

## 다음 턴(예고: S3 착수 전 빠른 스모크 테스트 스니펫)

기존 스크립트를 변경하지 않고, 새 추출기를 임시로 끼워 스모크만:

```python
# tools/smoke_extract.py  (임시 도우미; 커밋 제외 가능)
from pathlib import Path
from hvdc.extractors.case import extract_cases
from hvdc.extractors.site import extract_sites
from hvdc.extractors.lpo import extract_lpos

samples = [
    "RE: HVDC-ADOPT-HE-0476 - JPTW-71 / GRM-123",
    "PRL-O-046-O4(HE-0486) - PO 5001005009 AGI",
]

for s in samples:
    print(s)
    print("  cases:", [h["value"] for h in extract_cases(s)])
    print("  sites:", extract_sites(s))
    print("  lpos :", extract_lpos(s))
```

