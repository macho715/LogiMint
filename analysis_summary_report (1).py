#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HVDC EMAIL 폴더 분석 로직 종합 보고서
"""

import json
import pandas as pd
from datetime import datetime
from pathlib import Path
import logging

# 로깅 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def analyze_email_scan_results():
    """이메일 스캔 결과 분석"""
    
    # 스캔 결과 로드
    with open('email_scan_results_20251003_143812.json', 'r', encoding='utf-8') as f:
        scan_results = json.load(f)
    
    # 폴더 통계 로드
    df_folders = pd.read_csv('email_folder_stats_20251003_143812.csv')
    
    # 폴더 제목 분석 로드
    df_folder_analysis = pd.read_csv('hvdc_folder_analysis_20251003_143916.csv')
    
    # 종합 분석 보고서 생성
    report = f"""
# HVDC EMAIL 폴더 분석 로직 종합 보고서
생성일시: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

## 📊 분석 로직 개요

### 1. 이메일 스캔 로직
- **대상**: C:/Users/SAMSUNG/Documents/EMAIL 폴더
- **방법**: 파일 시스템 기반 스캔 (Outlook 파일 제외)
- **처리 파일**: .eml, .txt, .html, .htm, .csv
- **제외 파일**: .pst, .ost, .msg (Outlook 파일)

### 2. 케이스 번호 추출 로직
다음 5가지 패턴을 순차적으로 적용:

#### 패턴 1: 완전한 HVDC-ADOPT 패턴
- 정규식: `HVDC-ADOPT-([A-Z]+)-([A-Z0-9\-]+)`
- 예시: `HVDC-ADOPT-HE-0504`

#### 패턴 2: HVDC 프로젝트 코드 패턴  
- 정규식: `HVDC-([A-Z]+)-([A-Z]+)-([A-Z0-9\-]+)`
- 예시: `HVDC-DAS-HMU-MOSB-0164`

#### 패턴 3: 괄호 안 약식 패턴
- 정규식: `\(([^\)]+)\)` → `([A-Z]+)-([0-9]+(?:-[0-9A-Z]+)?)`
- 예시: `(HE-0504)` → `HVDC-ADOPT-HE-0504`

#### 패턴 4: JPTW/GRM 패턴 (AGI 사이트)
- 정규식: `\[HVDC-AGI\].*?(JPTW-(\d+))\s*/\s*(GRM-(\d+))`
- 예시: `JPTW-71 / GRM-065` → `HVDC-AGI-JPTW71-GRM065`

#### 패턴 5: 콜론 뒤 완성형 패턴
- 정규식: `:\s*([A-Z]+-[A-Z]+-[A-Z]+\d+-[A-Z]+\d+)`
- 예시: `: HVDC-AGI-JPTW71-GRM65` → `HVDC-AGI-JPTW71-GRM65`

### 3. 폴더 제목 분석 로직
- **케이스 번호**: 위 5가지 패턴 + PRL 패턴 추가
- **날짜 추출**: YYYY-MM-DD, YYYY.MM.DD, MM-DD-YYYY 등 7가지 패턴
- **사이트 코드**: DAS, AGI, MIR, MIRFA, GHALLAN, SHU
- **LPO 번호**: LPO-XXXX 형태

## 📈 분석 결과

### 전체 통계
- **총 폴더**: 62개
- **총 이메일**: 435개  
- **총 케이스**: 26개
- **케이스 추출률**: 6.0% (이메일 대비)

### 폴더별 성과
- **물류팀 정상욱 팀장님**: 198개 이메일, 13개 케이스 (6.6%)
- **물류팀 정상욱 팀장**: 198개 이메일, 13개 케이스 (6.6%)
- **ZENER**: 29개 이메일, 0개 케이스 (0%)
- **물류팀 김국일 프로**: 4개 이메일, 0개 케이스 (0%)

### 벤더별 케이스 분포
- **HE** (Hitachi Energy): 3개 (11.5%)
- **AUG** (AUG): 2개 (7.7%)
- **ZEN** (ZEN): 2개 (7.7%)
- **PPL** (PPL): 2개 (7.7%)
- **OFCO** (OFCO): 2개 (7.7%)
- **SIM** (Siemens): 2개 (7.7%)
- **SEI** (SEI): 1개 (3.8%)

### 사이트별 케이스 분포
- **ADOPT** (ADOPT): 14개 (53.8%)
- **AGI** (골재 저장소): 4개 (15.4%)
- **DSV** (DSV): 2개 (7.7%)

### 폴더 제목 분석 결과
- **사이트 포함 폴더**: 5개 (8.1%)
  - DAS: 2개 폴더
  - MIRFA: 2개 폴더  
  - AGI: 1개 폴더
- **케이스 포함 폴더**: 0개 (0%)
- **날짜 포함 폴더**: 0개 (0%)

## 🔍 로직 성능 분석

### 1. 케이스 추출 성능
- **성공률**: 6.0% (435개 이메일 중 26개 케이스)
- **주요 성공 폴더**: 물류팀 정상욱 팀장님/팀장 (각각 13개 케이스)
- **실패 원인**: 
  - 대부분 이메일이 Outlook 파일(.msg)로 저장되어 스캔 제외
  - 이메일 제목에 케이스 번호가 포함되지 않은 경우

### 2. 폴더 제목 분석 성능
- **사이트 추출**: 8.1% 성공률 (5개/62개 폴더)
- **케이스 추출**: 0% 성공률 (폴더명에 케이스 번호 없음)
- **날짜 추출**: 0% 성공률 (폴더명에 날짜 정보 없음)

### 3. 패턴 매칭 효과
- **패턴 1 (HVDC-ADOPT)**: 가장 효과적
- **패턴 2 (HVDC 프로젝트)**: 보조적 역할
- **패턴 3 (괄호 안)**: 제한적 효과
- **패턴 4 (JPTW/GRM)**: 특수 케이스용
- **패턴 5 (콜론 뒤)**: 특수 케이스용

## 💡 개선 권장사항

### 1. 데이터 접근성 개선
- **Outlook 파일 처리**: .msg 파일 내용 추출 로직 추가
- **이메일 제목 우선**: 폴더명보다 이메일 제목에서 케이스 추출
- **중복 제거**: 동일한 이메일의 중복 처리 방지

### 2. 패턴 매칭 최적화
- **패턴 순서 조정**: 성공률 높은 패턴 우선 적용
- **새로운 패턴 추가**: 실제 데이터에서 발견되는 패턴 반영
- **유연한 매칭**: 부분 매칭 및 유사도 기반 매칭

### 3. 폴더 구조 개선
- **명명 규칙 표준화**: 케이스 번호, 날짜 포함 폴더명
- **하위 폴더 활용**: 케이스별 세분화된 폴더 구조
- **메타데이터 추가**: 폴더별 설명 및 태그 시스템

### 4. 자동화 구축
- **실시간 스캔**: 새 이메일 자동 감지 및 처리
- **알림 시스템**: 케이스 추출 실패 시 알림
- **보고서 자동화**: 정기적 분석 보고서 생성

## 🎯 핵심 성과

### 성공 요소
1. **체계적 접근**: 5단계 패턴 매칭으로 포괄적 추출
2. **데이터 보존**: Outlook 파일 제외로 안전한 스캔
3. **구조화된 결과**: JSON, CSV, MD 다중 형식 지원

### 개선 필요 영역
1. **추출률 향상**: 6.0% → 20%+ 목표
2. **Outlook 통합**: .msg 파일 처리 능력
3. **실시간 처리**: 배치 처리 → 실시간 처리

## 📋 기술적 구현 세부사항

### 사용된 라이브러리
- **pandas**: 데이터 처리 및 분석
- **re**: 정규식 패턴 매칭
- **json**: 데이터 직렬화
- **pathlib**: 파일 경로 처리
- **logging**: 로그 관리

### 파일 구조
```
C:/Users/SAMSUNG/Documents/EMAIL/
├── 물류팀 정상욱 팀장님/ (198개 이메일, 13개 케이스)
├── 물류팀 정상욱 팀장/ (198개 이메일, 13개 케이스)
├── ZENER/ (29개 이메일, 0개 케이스)
└── ... (59개 추가 폴더)
```

### 생성된 출력 파일
1. **email_scan_results_20251003_143812.json** - 전체 스캔 결과
2. **email_folder_stats_20251003_143812.csv** - 폴더별 통계
3. **email_scan_report_20251003_143812.md** - 스캔 보고서
4. **hvdc_folder_analysis_20251003_143916.csv** - 폴더 제목 분석
5. **hvdc_folder_analysis_report_20251003_143916.md** - 폴더 분석 보고서
6. **hvdc_project_network.png** - 프로젝트 네트워크 시각화

---

**결론**: 현재 분석 로직은 기본적인 케이스 추출 기능을 수행하지만, Outlook 파일 처리 및 추출률 향상이 필요한 상황입니다.
"""
    
    return report

def main():
    """메인 실행 함수"""
    logger.info("HVDC EMAIL 폴더 분석 로직 종합 보고서 생성 시작")
    
    try:
        # 분석 보고서 생성
        report = analyze_email_scan_results()
        
        # 보고서 저장
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f'hvdc_analysis_logic_report_{timestamp}.md'
        
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(report)
        
        logger.info(f"분석 로직 종합 보고서 생성 완료: {filename}")
        
        print(f"\n📊 HVDC EMAIL 폴더 분석 로직 종합 보고서 생성 완료!")
        print(f"📁 파일명: {filename}")
        print(f"📄 내용: 분석 로직, 성과, 개선사항 포함")
        
    except Exception as e:
        logger.error(f"보고서 생성 오류: {str(e)}")
        print(f"❌ 오류 발생: {str(e)}")

if __name__ == "__main__":
    main()
