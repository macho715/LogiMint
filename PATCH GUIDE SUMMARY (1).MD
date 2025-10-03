# ExecSummary (요약)

* **모듈 분해**: `scanner`/`parser`/`patterns`/`extractors`/`report`/`cli` 계층으로 나누어 결합도↓ 응집도↑.
* **공통 유틸/예외/설정 통합**: `core/config.py`, `core/errors.py`, `core/regex.py`, `core/io.py`로 중복 제거.
* **케이스 추출 로직 표준화**: 각 스크립트에 흩어진 HVDC/PRL/JPTW/GRM 패턴을 단일 소스에서 재사용.
* **I/O 경계 분리**: 파일/폴더 스캔과 파싱, 가공, 리포팅을 분리해 테스트 가능한 순수 함수화.
* **테스트 포트폴리오**: 유닛(대다수) ≫ 통합 ≫ E2E(핵심 여정만). try/except 브랜치, 정규식, 경계 케이스 위주. 

---

# 리팩토링 타깃(현 상태 짚기)

* **패턴·추출이 흩어짐**

  * 케이스/사이트/LPO 등 추출 정규식이 각 파일에 중복·변형돼 존재. 예: `comprehensive_email_mapper.py`의 `self.case_patterns`와 추출기들, `folder_title_mapper.py`의 유사 패턴군.  
  * `create_complete_email_excel.py`·`update_email_pattern_rules.py`도 별도 구현.  

* **스캐닝/파싱/리포팅 결합**

  * `email_folder_scanner.py`가 파일 시스템 순회·읽기·정규식 추출을 한곳에서 처리. 테스트 어려움. 
  * 네트워크x/시각화 로직까지 섞인 곳도 존재(`comprehensive_email_mapper`). 

* **에러 처리 불균형**

  * 일부 스크립트 try/except 거의 없음(`create_complete_email_excel.py`, `update_email_pattern_rules.py`).  

* **실행 오케스트레이션 분산**

  * `run_all_scripts.py`(=HVDCScriptRunner) 존재하나 디펜던시·타임아웃만 관리. 파이프라인·계약(입출 규격) 계층화 필요. 

---

# 목표 아키텍처 (구조 커밋만, 행위 불변)

```
hvdc/
  core/
    config.py         # 경로/패턴/필드 스키마/로깅 설정
    errors.py         # DomainError, ParseError, IoError 등
    regex.py          # 표준 정규식 카탈로그(HVDC/PRL/JPTW/GRM/LPO/사이트)
    io.py             # 파일 읽기/쓰기/CSV/XLSX I/O 어댑터
    typing.py         # TypedDict/Protocol (Email, FolderMeta, CaseHit 등)
  scanner/
    fs_scanner.py     # 파일 시스템/확장자 필터/샘플 제한
    email_reader.py   # .eml/.txt/.html 헤더/본문 추출
  parser/
    subject_parser.py # 제목 → 메타데이터(case/site/lpo/phase) 추출 (순수함수)
    folder_parser.py  # 폴더명 → 메타데이터 추출 (순수함수)
  extractors/
    case.py           # 패턴 조합→케이스 리스트 반환 (중복 제거)
    site.py, lpo.py, phase.py ...
  report/
    timeline.py       # 타임라인/네트워크/집계 생성(순수 + pandas)
    excel.py          # 엑셀 아웃(기존 시트 구성 유지)
  cli/
    map_emails.py     # 기존 comprehensive_* 진입점 래핑
    scan_folders.py   # 기존 email_folder_scanner 래핑
    map_folders.py    # 기존 folder_title_mapper 래핑
    cargo_track.py    # hvdc_cargo_tracking_system 래핑
  pipeline/
    run_all.py        # run_all_scripts 대체: 단계/의존/계약 체크
tests/
  unit/ ...           # regex, parser, extractor, io 단위 테스트
  integration/ ...    # 스캐너→파서→리포트 흐름
  e2e/ ...            # 핵심 여정(“EMAIL 폴더→엑셀 5시트”)
```

* 근거: 현재 파일의 역할을 계층으로 분리 (스캔/파싱/추출/리포트/CLI)하며, 공통 패턴과 I/O는 `core`에 수렴.     

---

# 단계별 작업 (모두 **Structural** 커밋 / 테스트 그린 유지)

### S1) 공통 코어 신설

* `core/regex.py`에 **단일 정규식 카탈로그** 정의:

  * `HVDC_ADOPT = r'HVDC-ADOPT-([A-Z]+)-([A-Z0-9\-]+)'`
  * `HVDC_GENERIC = r'HVDC-([A-Z]+)-([A-Z]+)-([A-Z0-9\-]+)'`
  * `PAREN_VENDOR = r'\(([^\)]+)\)'` + `INNER = r'([A-Z]+)-([0-9]+(?:-[0-9A-Z]+)?)'`
  * `JPTW_GRM = r'\[HVDC-AGI\].*?(JPTW-(\d+))\s*/\s*(GRM-(\d+))'`
  * `TRAILING = r':\s*([A-Z]+-[A-Z]+-[A-Z]+\d+-[A-Z]+\d+)'`
  * `PRL = r'PRL-([A-Z]+)-(\d+)-([A-Z]+)-([A-Z0-9\-]+)'`
    *(위 패턴은 현 스크립트들에서 사용되는 변형 포함—중복원 제거)*    
* `core/errors.py`: `ParseError`, `PatternError`, `ScanError` 등 도메인 예외.
* `core/config.py`: 이메일 루트 경로, 확장자 화이트리스트, 샘플 제한 수, 출력 파일 네이밍 규칙. (기존 하드코딩 분리) 
* `core/io.py`: 파일 읽기/쓰기·로깅 초기화·timestamp 유틸.

### S2) 추출기 통합

* `extractors/case.py`에 **단일 `extract_cases(text:str)->list[str]`** 구현. 기존 각 파일의 `extract_case_numbers*` 대체(동일 결과).    
* `extractors/site.py`, `lpo.py`, `phase.py`도 동일 구조의 순수함수.

### S3) 파서·스캐너 분리

* `scanner/fs_scanner.py`: 폴더 순회·확장자 필터·Outlook 제외(기존 로직 유지). `scan_folder(path)->Iterable[Path]`. 
* `parser/subject_parser.py`/`folder_parser.py`: 텍스트/폴더명→(cases, sites, lpos, phase) dict 반환. 기존 클래스의 메서드를 **순수 함수**로 추출.  

### S4) 리포트 출력 계층화

* `report/excel.py`: 기존 5시트 구성(전체/사이트별/LPO/긴급/배송) 유지. `DataFrame→Excel`만 담당. (현재 시트 구성 로직 `create_complete_email_excel.py`에서 추출) 
* `report/timeline.py`: timeline/네트워크 집계만 담당(현재 `comprehensive_email_mapper` 섞임 분리). 

### S5) CLI 래퍼로 하위호환

* 기존 스크립트 이름은 유지하되 내부에서 새 모듈 호출(동작 동일). `map_emails.py` 등.
* `pipeline/run_all.py`는 `run_all_scripts.py` 기능+의존성 선언/결과 요약만 담당. 

> 위 5스텝은 **행위 불변/리팩토링 전용 커밋(예: `structural(core): extract regex catalog`)**으로 나갑니다. 테스트 전부 그린 상태로만 전개. 

---

# 테스트 전략 (생성할 테스트 예시)

### 유닛 (주력)

* `tests/unit/test_regex_cases.py`

  * HVDC/PRL/JPTW·GRM/괄호 패턴 파라메트릭 테스트(대표 제목 20개).
* `tests/unit/test_subject_parser.py`

  * “Title→cases/sites/lpos/phase” 결정 테이블 기반.
* `tests/unit/test_folder_parser.py`

  * 폴더명 날짜/사이트/PRL 패턴 추출.
* `tests/unit/test_io_guard.py`

  * `.pst/.ost/.msg` 제외, `.eml/.txt/.html/.csv` 허용. (기존 스캐너 규칙) 

### 통합

* “폴더 샘플 → 스캔 → 파싱 → DataFrame” 파이프 테스트(디스크 의존 최소화, 임시폴더).

### E2E (핵심 여정만)

* “EMAIL 폴더 → 엑셀 5시트 생성” 1케이스, 결과 행 수·필드 존재성 검증. (느리면 `@ignore` 태깅) 

---

# 에러 처리 표준 (코드 가이드)

* 경계 I/O만 try/except. 내부 순수함수는 예외 전파.
* 예:

```python
# core/errors.py
class ScanError(Exception): ...
class ParseError(Exception): ...
```

```python
# scanner/fs_scanner.py
def scan_folder(root: Path) -> list[Path]:
    try:
        # walk & filter
        return files
    except OSError as e:
        raise ScanError(f"scan failed: {root}") from e
```

현재 try/except 부족한 스크립트는 위 규칙으로 보완 (엑셀 생성/파일 읽기/정규식 실패).  

---

# 마이그레이션 매핑 (구 파일 → 신 모듈)

| 기존                               | 새 배치                                                                                  | 비고                  |
| -------------------------------- | ------------------------------------------------------------------------------------- | ------------------- |
| `comprehensive_email_mapper.py`  | `parser/subject_parser.py`, `extractors/*`, `report/timeline.py`, `cli/map_emails.py` | 패턴·시각화 분리           |
| `folder_title_mapper.py`         | `parser/folder_parser.py`, `extractors/*`, `cli/map_folders.py`                       | 패턴 공유화              |
| `email_folder_scanner.py`        | `scanner/fs_scanner.py`                                                               | I/O 경계화             |
| `create_complete_email_excel.py` | `report/excel.py`                                                                     | 시트 생성 전용화           |
| `update_email_pattern_rules.py`  | `extractors/case.py` + 테스트                                                            | 규칙 업데이트는 테스트 주도     |
| `hvdc_cargo_tracking_system.py`  | `cli/cargo_track.py`(래핑)                                                              | 내부 로직은 후속(행위 변경 시)  |
| `run_all_scripts.py`             | `pipeline/run_all.py`                                                                 | 의존·결과 집계 개선         |

---

# 커밋 시퀀스 (예시)

1. `structural(core): add regex/typing/errors/config/io scaffolding`
2. `structural(extractors): unify case/site/lpo/phase extractors`
3. `structural(parser): extract subject_parser & folder_parser (pure)`
4. `structural(scanner): move fs scan & filters to scanner/fs_scanner.py`
5. `structural(report): split excel/timeline generators`
6. `structural(cli): thin wrappers to keep original entrypoints`
7. `tests(unit): add regex/parse/io unit tests (RED→GREEN)`
8. `tests(integration): add scan→parse pipeline test`
9. `tests(e2e): add EMAIL→Excel minimal journey (@ignore default)`

> 모든 커밋은 **행위 불변**을 유지(기존 출력 대비 샘플 비교). 기능개선은 이후 **behavioral:** 커밋 라인에서. 

---

# 리스크 & 완화

* 정규식 변경으로 인한 추출 차이 → **골든 샘플 제목** 세트로 스냅샷 테스트.
* 엑셀 시트 구성 차이 → 열 순서/시트명 스냅샷 비교(허용 diff는 타임스탬프만). 
* 파이프 분리로 생기는 경로/인코딩 이슈 → `core/io.py`에서 단일 처리.

---

# 바로 적용 체크리스트

* [ ] `core`/`extractors`/`parser`/`scanner`/`report`/`cli` 디렉터리 생성
* [ ] 공통 정규식 카탈로그 1원화 및 기존 호출 교체
* [ ] try/except 경계 I/O에 한정 적용
* [ ] 유닛 테스트(정규식/파서) 추가 후 전부 그린
* [ ] 기존 엔트리 스크립트는 CLI 래퍼로 유지(호환성)

---

