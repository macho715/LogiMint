# 🚀 LogiMint - HVDC Project Logistics Intelligence System

[![Python](https://img.shields.io/badge/Python-3.11+-blue.svg)](https://python.org)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Status](https://img.shields.io/badge/Status-Active-brightgreen.svg)](https://github.com/macho715/LogiMint)

> **Advanced Logistics Intelligence System for Samsung C&T HVDC Project**  
> 물류 자동화, 이메일 분석, 화물 추적을 위한 통합 플랫폼

## 📋 프로젝트 개요

LogiMint는 Samsung C&T HVDC 프로젝트를 위한 종합 물류 지능 시스템입니다. 이메일 자동 처리, 폴더 분석, 화물 추적 등 10개의 핵심 스크립트를 통합하여 물류 운영의 효율성을 극대화합니다.

### 🎯 주요 기능

- **📧 이메일 자동 처리**: Outlook 파일 제외 이메일 데이터 추출 및 분석
- **📁 폴더 지능 분석**: 제목 기반 케이스 번호, 날짜, 사이트 자동 매핑
- **🚢 화물 추적 시스템**: 실시간 화물 위치 및 상태 모니터링
- **📊 종합 보고서**: Excel 기반 분석 보고서 자동 생성
- **🔍 패턴 인식**: 정규식 기반 고급 패턴 매칭

## 🏗️ 시스템 아키텍처

```
LogiMint/
├── 📧 Email Processing
│   ├── email_folder_scanner.py      # 폴더 스캔 및 이메일 추출
│   ├── email_ontology_mapper.py     # 온톨로지 매핑
│   └── create_complete_email_excel.py # Excel 보고서 생성
├── 📁 Folder Analysis
│   ├── folder_title_mapper.py       # 제목 기반 매핑
│   └── simple_folder_analyzer.py    # 간단 분석
├── 🚢 Cargo Tracking
│   └── hvdc_cargo_tracking_system.py # 화물 추적 시스템
├── 📊 Comprehensive Analysis
│   ├── comprehensive_email_mapper.py # 종합 매핑
│   └── analysis_summary_report.py   # 분석 보고서
├── ⚙️ Utilities
│   ├── update_email_pattern_rules.py # 패턴 규칙 업데이트
│   └── monitor_scan_progress.py     # 진행상황 모니터링
└── 🧪 HVDC Core Module
    ├── core/                        # 핵심 기능
    ├── extractors/                  # 데이터 추출기
    ├── parser/                      # 파서 모듈
    ├── scanner/                     # 스캐너 모듈
    ├── report/                      # 보고서 생성
    └── tests/                       # 단위 테스트
```

## 🚀 빠른 시작

### 1. 저장소 클론
```bash
git clone https://github.com/macho715/LogiMint.git
cd LogiMint
```

### 2. 의존성 설치
```bash
pip install -r requirements.txt
```

### 3. 기본 실행
```bash
# 종합 이메일 요약 (JSON + JSONL 출력)
python comprehensive_email_mapper.py --email-root ./emails --output out/summary.json

# 폴더 제목 매핑
python folder_title_mapper.py --folder-root ./emails --output out/folder_mapping.json

# 화물 추적 요약 (상태 엑셀 필요)
python hvdc_cargo_tracking_system.py --status-file data/status.xlsx --output out/cargo_summary.json
```

### 4. 환경 설정
- `.env` 또는 환경 변수로 다음 값을 오버라이드할 수 있습니다.
  - `HVDC_EMAIL_ROOT`: 기본 이메일 루트 경로
  - `HVDC_LOG_DIR`: 로그 파일 저장 위치
  - `HVDC_LOG_JSON`: `1` 설정 시 JSON 라인 로깅 사용
  - `HVDC_ASYNC_BATCH_SIZE`: 이메일 파싱 비동기 배치 크기
  - `HVDC_EXCEL_CHUNK_SIZE`: 엑셀 쓰기 시 청크 크기
- 모든 경로/토글은 `config/settings.py`에서 중앙 관리됩니다.

## 📊 스크립트 상세 정보

| 스크립트 | 크기 | 라인 수 | 기능 | 복잡도 |
|---------|------|---------|------|--------|
| **comprehensive_email_mapper.py** | 8KB | 121 | 비동기 이메일 요약 매퍼 | 🟡 중간 |
| **folder_title_mapper.py** | 4KB | 64 | 폴더 제목 기반 매핑 | 🟢 단순 |
| **hvdc_cargo_tracking_system.py** | 12KB | 172 | 화물 추적 시스템 | 🟡 중간 |
| **email_ontology_mapper.py** | 19KB | 440 | 온톨로지 매핑 | 🟡 중간 |
| **email_folder_scanner.py** | 17KB | 434 | 폴더 스캔 | 🟡 중간 |
| **create_complete_email_excel.py** | 15KB | 336 | Excel 보고서 생성 | 🟡 중간 |
| **simple_folder_analyzer.py** | 11KB | 296 | 간단 폴더 분석 | 🟡 중간 |
| **update_email_pattern_rules.py** | 9KB | 254 | 패턴 규칙 업데이트 | 🟡 중간 |
| **analysis_summary_report.py** | 8KB | 222 | 분석 보고서 | 🟡 중간 |
| **monitor_scan_progress.py** | 2KB | 71 | 진행상황 모니터링 | 🟢 단순 |

## 🔧 사용 방법

### 기본 워크플로우
1. **데이터 수집**: `utils/file_handler.py` 기반 스캐너로 이메일 경로를 수집합니다.
2. **패턴 매핑**: `utils/pattern_matcher.py`가 케이스/벤더/사이트를 정규화합니다.
3. **결과 요약**: `comprehensive_email_mapper.py` 비동기 파이프라인이 JSON/JSONL을 생성합니다.
4. **폴더 맵**: `folder_title_mapper.py`가 폴더-케이스 매핑 JSON을 저장합니다.
5. **화물 요약**: `hvdc_cargo_tracking_system.py`가 상태 엑셀을 집계해 요약 JSON을 생성합니다.

### 고급 사용법
```bash
# 특정 패턴 업데이트
python update_email_pattern_rules.py

# 화물 추적 시스템 실행
python hvdc_cargo_tracking_system.py

# 전체 파이프라인 실행
python run_all_scripts.py
```

## 🧪 테스트

### 단위 테스트 및 품질 게이트
```bash
# 전체 테스트 및 커버리지
pytest -q
coverage run -m pytest && coverage report

# 코드 품질 게이트
black --check .
isort --check-only .
flake8 .
mypy --strict .
```

## 📈 성능 지표

### 현재 상태
- **총 코드 라인**: 3,530 lines
- **총 파일 크기**: 158KB
- **함수 수**: 76개
- **에러 처리**: 15개 try-except 블록

### 목표 지표
- **코드 복잡도**: 평균 라인 수 < 300
- **에러 처리**: try-except 비율 > 5%
- **테스트 커버리지**: > 80%
- **실행 시간**: 기존 대비 50% 단축

## 🔄 개발 로드맵

### Phase 1: 구조 개선 (진행 중)
- [x] GitHub 저장소 설정
- [x] 기본 문서화
- [ ] 공통 모듈 분리
- [ ] 설정 관리 통합
- [ ] 로깅 표준화

### Phase 2: 코드 품질 개선
- [ ] 우선순위 스크립트 리팩토링
- [ ] 에러 처리 표준화
- [ ] 단위 테스트 추가
- [ ] TDD 방법론 적용

### Phase 3: 성능 및 기능 개선
- [ ] 비동기 처리 도입
- [ ] 메모리 최적화
- [ ] 고급 분석 기능
- [ ] API 인터페이스 구축

## 🤝 기여하기

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

### 코딩 스타일
- **Python**: PEP 8 준수
- **커밋 메시지**: [Structural] 또는 [Behavioral] 접두사 사용
- **테스트**: 모든 새 기능에 대한 테스트 코드 작성 필수

## 📝 라이선스

이 프로젝트는 MIT 라이선스 하에 배포됩니다. 자세한 내용은 [LICENSE](LICENSE) 파일을 참조하세요.

## 📞 연락처

- **프로젝트**: Samsung C&T HVDC Project
- **GitHub**: [macho715](https://github.com/macho715)
- **저장소**: [LogiMint](https://github.com/macho715/LogiMint)

## 🙏 감사의 말

- Samsung C&T 물류팀
- HVDC 프로젝트 개발팀
- 오픈소스 커뮤니티

---

**LogiMint** - 물류 지능의 새로운 표준을 제시합니다. 🚀

*마지막 업데이트: 2025-01-03*
