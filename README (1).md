# HVDC 프로젝트 주요 스크립트 통합 폴더

## 📁 폴더 개요
이 폴더는 HVDC 프로젝트의 주요 10개 스크립트를 한 곳에 모아놓은 통합 폴더입니다.

## 📋 포함된 스크립트 목록

### 1. 이메일 처리 스크립트
- **`email_folder_scanner.py`** (17,562 bytes, 434 lines)
  - EMAIL 폴더의 모든 하위 폴더 스캔
  - Outlook 파일 제외하고 이메일 데이터 추출
  - 12개 함수, 4개 try-except 블록

- **`email_ontology_mapper.py`** (19,276 bytes, 440 lines)
  - 이메일 데이터를 HVDC 온톨로지 시스템에 매핑
  - 10개 함수, 2개 try-except 블록

- **`create_complete_email_excel.py`** (15,074 bytes, 336 lines)
  - 매핑된 이메일 데이터를 종합 엑셀 보고서로 생성
  - 5개 함수, 에러 처리 없음

### 2. 폴더 분석 스크립트
- **`folder_title_mapper.py`** (21,568 bytes, 524 lines)
  - 폴더 제목 기반 케이스 번호, 날짜, 사이트 매핑
  - 10개 함수, 1개 try-except 블록

- **`simple_folder_analyzer.py`** (11,866 bytes, 296 lines)
  - 폴더 제목 간단 분석
  - 9개 함수, 1개 try-except 블록

### 3. 종합 매핑 스크립트
- **`comprehensive_email_mapper.py`** (23,027 bytes, 556 lines)
  - 종합 이메일 매핑 및 네트워크 시각화
  - 11개 함수, 2개 try-except 블록

### 4. 화물 추적 스크립트
- **`hvdc_cargo_tracking_system.py`** (20,074 bytes, 497 lines)
  - HVDC 화물 추적 시스템
  - 11개 함수, 2개 try-except 블록

### 5. 패턴 업데이트 스크립트
- **`update_email_pattern_rules.py`** (9,357 bytes, 254 lines)
  - 이메일 패턴 규칙 업데이트
  - 5개 함수, 에러 처리 없음

### 6. 모니터링 스크립트
- **`monitor_scan_progress.py`** (2,235 bytes, 71 lines)
  - 스캔 진행상황 모니터링
  - 1개 함수, 에러 처리 없음

### 7. 분석 보고서 스크립트
- **`analysis_summary_report.py`** (8,123 bytes, 222 lines)
  - 분석 로직 종합 보고서 생성
  - 2개 함수, 1개 try-except 블록

## 🔧 사용 방법

### 기본 실행 순서
1. **폴더 스캔**: `python email_folder_scanner.py`
2. **진행상황 모니터링**: `python monitor_scan_progress.py`
3. **이메일 매핑**: `python email_ontology_mapper.py`
4. **엑셀 보고서 생성**: `python create_complete_email_excel.py`
5. **종합 분석**: `python comprehensive_email_mapper.py`

### 개별 실행
각 스크립트는 독립적으로 실행 가능합니다.

## 📊 스크립트 통계

| 스크립트명 | 크기 | 라인 수 | 함수 수 | 에러 처리 | 복잡도 |
|------------|------|---------|---------|-----------|--------|
| comprehensive_email_mapper.py | 23KB | 556 | 11 | 2 | 🔴 복잡 |
| folder_title_mapper.py | 21KB | 524 | 10 | 1 | 🔴 복잡 |
| hvdc_cargo_tracking_system.py | 20KB | 497 | 11 | 2 | 🟡 중간 |
| email_ontology_mapper.py | 19KB | 440 | 10 | 2 | 🟡 중간 |
| email_folder_scanner.py | 17KB | 434 | 12 | 4 | 🟡 중간 |
| create_complete_email_excel.py | 15KB | 336 | 5 | 0 | 🟡 중간 |
| simple_folder_analyzer.py | 11KB | 296 | 9 | 1 | 🟡 중간 |
| update_email_pattern_rules.py | 9KB | 254 | 5 | 0 | 🟡 중간 |
| analysis_summary_report.py | 8KB | 222 | 2 | 1 | 🟡 중간 |
| monitor_scan_progress.py | 2KB | 71 | 1 | 0 | 🟢 단순 |

## 🚨 주요 문제점

### 1. 코드 품질 문제
- **복잡한 스크립트**: 2개 스크립트가 500라인 이상
- **에러 처리 부족**: 3개 스크립트에 try-except 블록 없음
- **에러 처리 비율 낮음**: 평균 0.5% (권장: 5% 이상)

### 2. 구조적 문제
- **단일 책임 원칙 위반**: 하나의 스크립트가 너무 많은 기능 담당
- **코드 중복**: 유사한 로직이 여러 스크립트에 반복
- **하드코딩**: 설정값이 코드에 직접 포함

## 💡 개선 방안

### Phase 1: 기반 구조 개선
1. **공통 모듈 분리**
   - `utils/email_parser.py`: 이메일 파싱 공통 로직
   - `utils/pattern_matcher.py`: 패턴 매칭 공통 로직
   - `utils/file_handler.py`: 파일 처리 공통 로직

2. **설정 관리 통합**
   - `config/settings.py`: 모든 설정값 중앙 관리
   - `config/patterns.py`: 정규식 패턴 정의

### Phase 2: 코드 품질 개선
1. **리팩토링 우선순위**
   - `comprehensive_email_mapper.py` (556 lines)
   - `folder_title_mapper.py` (524 lines)
   - `hvdc_cargo_tracking_system.py` (497 lines)

2. **에러 처리 표준화**
   - 모든 스크립트에 try-except 블록 추가
   - 커스텀 예외 클래스 정의

### Phase 3: 성능 및 안정성 개선
1. **비동기 처리 도입**
2. **메모리 최적화**
3. **테스트 코드 추가**

## 📋 실행 체크리스트

### 즉시 실행 가능
- [ ] 공통 모듈 디렉토리 생성
- [ ] 설정 파일 통합
- [ ] 로깅 표준화

### 단기 개선 (1-2주)
- [ ] 우선순위 스크립트 리팩토링
- [ ] 에러 처리 표준화
- [ ] 단위 테스트 추가

### 중기 개선 (1-2개월)
- [ ] 비동기 처리 도입
- [ ] 성능 최적화
- [ ] 통합 테스트 구축

## 🎯 성공 지표

### 코드 품질
- **코드 복잡도**: 평균 라인 수 < 300
- **에러 처리**: try-except 비율 > 5%
- **테스트 커버리지**: > 80%

### 성능
- **실행 시간**: 기존 대비 50% 단축
- **메모리 사용량**: 기존 대비 30% 감소
- **에러율**: < 1%

### 유지보수성
- **코드 중복**: < 10%
- **의존성**: 최소화 및 명확화
- **문서화**: 모든 함수에 docstring

---

**생성일시**: 2025-10-03 14:46:00  
**총 스크립트 수**: 10개  
**총 코드 라인**: 3,530 lines  
**총 파일 크기**: 158,158 bytes
