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
# 폴더 스캔 및 이메일 분석
python email_folder_scanner.py

# 진행상황 모니터링
python monitor_scan_progress.py

# 종합 분석 실행
python comprehensive_email_mapper.py
```

## 📊 스크립트 상세 정보

| 스크립트 | 크기 | 라인 수 | 기능 | 복잡도 |
|---------|------|---------|------|--------|
| **comprehensive_email_mapper.py** | 23KB | 556 | 종합 이메일 매핑 및 네트워크 시각화 | 🔴 복잡 |
| **folder_title_mapper.py** | 21KB | 524 | 폴더 제목 기반 매핑 | 🔴 복잡 |
| **hvdc_cargo_tracking_system.py** | 20KB | 497 | 화물 추적 시스템 | 🟡 중간 |
| **email_ontology_mapper.py** | 19KB | 440 | 온톨로지 매핑 | 🟡 중간 |
| **email_folder_scanner.py** | 17KB | 434 | 폴더 스캔 | 🟡 중간 |
| **create_complete_email_excel.py** | 15KB | 336 | Excel 보고서 생성 | 🟡 중간 |
| **simple_folder_analyzer.py** | 11KB | 296 | 간단 폴더 분석 | 🟡 중간 |
| **update_email_pattern_rules.py** | 9KB | 254 | 패턴 규칙 업데이트 | 🟡 중간 |
| **analysis_summary_report.py** | 8KB | 222 | 분석 보고서 | 🟡 중간 |
| **monitor_scan_progress.py** | 2KB | 71 | 진행상황 모니터링 | 🟢 단순 |

## 🔧 사용 방법

### 기본 워크플로우
1. **데이터 수집**: `email_folder_scanner.py`로 이메일 데이터 추출
2. **진행 모니터링**: `monitor_scan_progress.py`로 상태 확인
3. **데이터 매핑**: `email_ontology_mapper.py`로 온톨로지 매핑
4. **보고서 생성**: `create_complete_email_excel.py`로 Excel 보고서 생성
5. **종합 분석**: `comprehensive_email_mapper.py`로 네트워크 분석

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

### 단위 테스트 실행
```bash
# 전체 테스트 실행
python -m pytest tests/

# 특정 모듈 테스트
python -m pytest tests/unit/test_regex_cases.py

# 커버리지 포함 테스트
python -m pytest --cov=hvdc tests/
```

### 스모크 테스트
```bash
# 기본 기능 검증
python tools/smoke_extract.py
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
