#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HVDC 폴더 제목 기반 매핑 시스템
폴더 제목에서 케이스 번호, 날짜, 프로젝트 정보를 추출하여 매핑
"""

import os
import re
import json
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path
import logging
from typing import Dict, List, Tuple, Optional
from collections import defaultdict, Counter

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('folder_mapping.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class FolderTitleMapper:
    """폴더 제목 기반 매핑 클래스"""
    
    def __init__(self, email_root_path: str):
        self.email_root_path = Path(email_root_path)
        self.folder_mappings = []
        self.case_network = defaultdict(list)
        self.date_network = defaultdict(list)
        self.project_network = defaultdict(list)
        
        # HVDC 케이스 번호 추출 패턴
        self.case_patterns = {
            'pattern1': r'HVDC-ADOPT-([A-Z]+)-([A-Z0-9\-]+)',
            'pattern2': r'HVDC-([A-Z]+)-([A-Z]+)-([A-Z0-9\-]+)',
            'pattern3_outer': r'\(([^\)]+)\)',
            'pattern3_inner': r'([A-Z]+)-([0-9]+(?:-[0-9A-Z]+)?)',
            'pattern4': r'\[HVDC-AGI\].*?(JPTW-(\d+))\s*/\s*(GRM-(\d+))',
            'pattern5': r':\s*([A-Z]+-[A-Z]+-[A-Z]+\d+-[A-Z]+\d+)',
            'pattern6': r'PRL-([A-Z]+)-(\d+)-([A-Z]+)-([A-Z0-9\-]+)',
            'pattern7': r'\(([A-Z]+)-(\d+)\)'
        }
        
        # 벤더 코드 매핑
        self.vendor_codes = {
            'HE': 'Hitachi Energy',
            'SCT': 'Samsung C&T', 
            'SIM': 'Siemens',
            'MOSB': 'MOSB',
            'ALS': 'ALS',
            'JPTW': 'JPTW',
            'GRM': 'GRM',
            'ZEN': 'ZENER',
            'DAS': 'DAS Site',
            'AGI': 'AGI Site',
            'MIR': 'MIR Site',
            'MIRFA': 'MIRFA Site',
            'GHALLAN': 'GHALLAN Site',
            'SHU': 'Shuweihat',
            'BWE': 'Best Way Equipment',
            'FAL': 'Falcor Engineering',
            'HEC': 'Hanlim Engineering',
            'SPE': 'Super Phoenix',
            'NAF': 'NAF'
        }
        
        # 사이트 코드 매핑
        self.site_codes = {
            'DAS': '변전소',
            'AGI': '골재 저장소', 
            'MIR': '중간 변전소',
            'MIRFA': '보조 변전소',
            'GHALLAN': '사막 사이트',
            'SHU': 'Shuweihat 사이트'
        }
        
        # 프로젝트 단계별 키워드
        self.project_phases = {
            'procurement': ['LPO', 'PO', 'Purchase Order', 'Procurement', 'Order'],
            'shipping': ['Shipping', 'Delivery', 'Container', 'CNTR', 'LCT', 'Vessel', 'Cargo'],
            'customs': ['Customs', 'Clearance', 'Import', 'Export', 'Duty'],
            'logistics': ['Logistics', 'Transport', 'Freight', 'Material', 'Backload'],
            'installation': ['Installation', 'Install', 'Mounting', 'Assembly'],
            'testing': ['Test', 'Testing', 'Commissioning', 'Startup'],
            'certification': ['Certificate', 'Cert', 'MTC', 'COC', 'Quality']
        }

    def extract_case_numbers_from_folder(self, folder_name: str) -> List[str]:
        """폴더명에서 HVDC 케이스 번호 추출"""
        if not folder_name:
            return []
        
        case_numbers = []
        folder_str = str(folder_name).upper()
        
        # 패턴 1: 완전한 HVDC-ADOPT 패턴
        pattern1 = re.findall(self.case_patterns['pattern1'], folder_str, re.IGNORECASE)
        for match in pattern1:
            vendor_code = match[0].upper()
            case_num = match[1]
            full_case = f"HVDC-ADOPT-{vendor_code}-{case_num}"
            if full_case not in case_numbers:
                case_numbers.append(full_case)
        
        # 패턴 2: HVDC 프로젝트 코드 패턴
        pattern2 = re.findall(self.case_patterns['pattern2'], folder_str, re.IGNORECASE)
        for match in pattern2:
            site_code = match[0].upper()
            vendor_code = match[1].upper()
            case_num = match[2]
            full_case = f"HVDC-{site_code}-{vendor_code}-{case_num}"
            if full_case not in case_numbers:
                case_numbers.append(full_case)
        
        # 패턴 3: 괄호 안 약식 패턴 (HE-XXXX)
        pattern3_outer = re.findall(self.case_patterns['pattern3_outer'], folder_str)
        for outer_match in pattern3_outer:
            pattern3_inner = re.findall(self.case_patterns['pattern3_inner'], outer_match, re.IGNORECASE)
            for match in pattern3_inner:
                vendor_code = match[0].upper()
                case_num = match[1]
                full_case = f"HVDC-ADOPT-{vendor_code}-{case_num}"
                if full_case not in case_numbers:
                    case_numbers.append(full_case)
        
        # 패턴 4: JPTW/GRM 패턴
        pattern4 = re.findall(self.case_patterns['pattern4'], folder_str, re.IGNORECASE)
        for match in pattern4:
            jptw_num = match[1]
            grm_num = match[3]
            full_case = f"HVDC-AGI-JPTW{jptw_num}-GRM{grm_num}".upper()
            if full_case not in case_numbers:
                case_numbers.append(full_case)
        
        # 패턴 5: 콜론 뒤 완성형 패턴
        pattern5 = re.findall(self.case_patterns['pattern5'], folder_str, re.IGNORECASE)
        for match in pattern5:
            clean_case = re.sub(r'\(.*?\)', '', match).strip().upper()
            if clean_case not in case_numbers:
                case_numbers.append(clean_case)
        
        # 패턴 6: PRL 패턴 (PRL-SHU-053-O-INDOOR-SPARE)
        pattern6 = re.findall(self.case_patterns['pattern6'], folder_str, re.IGNORECASE)
        for match in pattern6:
            site_code = match[0].upper()
            prl_num = match[1]
            prl_type = match[2].upper()
            prl_desc = match[3]
            full_case = f"PRL-{site_code}-{prl_num}-{prl_type}-{prl_desc}"
            if full_case not in case_numbers:
                case_numbers.append(full_case)
        
        # 패턴 7: 괄호 안 단순 패턴 (HE-XXXX)
        pattern7 = re.findall(self.case_patterns['pattern7'], folder_str, re.IGNORECASE)
        for match in pattern7:
            vendor_code = match[0].upper()
            case_num = match[1]
            full_case = f"HVDC-ADOPT-{vendor_code}-{case_num}"
            if full_case not in case_numbers:
                case_numbers.append(full_case)
        
        return case_numbers

    def extract_dates_from_folder(self, folder_name: str) -> List[str]:
        """폴더명에서 날짜 추출"""
        if not folder_name:
            return []
        
        dates = []
        folder_str = str(folder_name)
        
        # 다양한 날짜 패턴
        date_patterns = [
            r'(\d{4}-\d{2}-\d{2})',  # YYYY-MM-DD
            r'(\d{4}\.\d{2}\.\d{2})',  # YYYY.MM.DD
            r'(\d{2}-\d{2}-\d{4})',  # MM-DD-YYYY
            r'(\d{2}/\d{2}/\d{4})',  # MM/DD/YYYY
            r'(\d{4}년\s*\d{1,2}월\s*\d{1,2}일)',  # YYYY년 MM월 DD일
            r'(\d{1,2}월\s*\d{1,2}일)',  # MM월 DD일
            r'(\d{4}-\d{2}-\d{2}\s*오[전후]\s*\d{1,2}_\d{2}_\d{2})',  # YYYY-MM-DD 오전/오후 HH_MM_SS
        ]
        
        for pattern in date_patterns:
            matches = re.findall(pattern, folder_str)
            for match in matches:
                if match not in dates:
                    dates.append(match)
        
        return dates

    def extract_project_phase_from_folder(self, folder_name: str) -> str:
        """폴더명에서 프로젝트 단계 추출"""
        folder_lower = str(folder_name).lower()
        
        for phase, keywords in self.project_phases.items():
            for keyword in keywords:
                if keyword.lower() in folder_lower:
                    return phase
        
        return 'general'

    def extract_lpo_numbers_from_folder(self, folder_name: str) -> List[str]:
        """폴더명에서 LPO 번호 추출"""
        lpo_pattern = r'LPO[-\s]?(\d+)'
        matches = re.findall(lpo_pattern, str(folder_name), re.IGNORECASE)
        return [f"LPO-{m}" for m in matches]

    def extract_site_codes_from_folder(self, folder_name: str) -> List[str]:
        """폴더명에서 사이트 코드 추출"""
        site_pattern = r'\b(DAS|AGI|MIR|MIRFA|GHALLAN|SHU)\b'
        matches = re.findall(site_pattern, str(folder_name), re.IGNORECASE)
        return [match.upper() for match in matches]

    def scan_all_folders(self):
        """모든 폴더 스캔 및 매핑"""
        logger.info(f"폴더 제목 기반 매핑 시작: {self.email_root_path}")
        
        if not self.email_root_path.exists():
            logger.error(f"EMAIL 폴더가 존재하지 않습니다: {self.email_root_path}")
            return
        
        folder_count = 0
        total_cases = 0
        total_dates = 0
        
        for folder_path in self.email_root_path.iterdir():
            if folder_path.is_dir():
                folder_count += 1
                folder_name = folder_path.name
                
                logger.info(f"폴더 분석 중: {folder_name}")
                
                # 케이스 번호 추출
                cases = self.extract_case_numbers_from_folder(folder_name)
                
                # 날짜 추출
                dates = self.extract_dates_from_folder(folder_name)
                
                # 프로젝트 단계 추출
                phase = self.extract_project_phase_from_folder(folder_name)
                
                # LPO 번호 추출
                lpos = self.extract_lpo_numbers_from_folder(folder_name)
                
                # 사이트 코드 추출
                sites = self.extract_site_codes_from_folder(folder_name)
                
                # 폴더 매핑 데이터 생성
                folder_mapping = {
                    'folder_name': folder_name,
                    'folder_path': str(folder_path),
                    'cases': cases,
                    'dates': dates,
                    'phase': phase,
                    'lpos': lpos,
                    'sites': sites,
                    'case_count': len(cases),
                    'date_count': len(dates),
                    'last_modified': datetime.fromtimestamp(folder_path.stat().st_mtime).isoformat()
                }
                
                self.folder_mappings.append(folder_mapping)
                
                # 네트워크 데이터 구축
                for case in cases:
                    self.case_network[case].append({
                        'folder': folder_name,
                        'phase': phase,
                        'dates': dates,
                        'sites': sites
                    })
                
                for date in dates:
                    self.date_network[date].append({
                        'folder': folder_name,
                        'cases': cases,
                        'phase': phase
                    })
                
                for site in sites:
                    self.project_network[site].append({
                        'folder': folder_name,
                        'cases': cases,
                        'dates': dates,
                        'phase': phase
                    })
                
                total_cases += len(cases)
                total_dates += len(dates)
        
        logger.info(f"폴더 스캔 완료: {folder_count}개 폴더, {total_cases}개 케이스, {total_dates}개 날짜")

    def generate_comprehensive_report(self) -> str:
        """종합 매핑 보고서 생성"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # 기본 통계
        total_folders = len(self.folder_mappings)
        total_cases = len(self.case_network)
        total_dates = len(self.date_network)
        total_sites = len(self.project_network)
        
        # 케이스별 통계
        case_stats = {}
        for case, folders in self.case_network.items():
            case_stats[case] = {
                'folder_count': len(folders),
                'phases': list(set([f['phase'] for f in folders])),
                'sites': list(set([site for f in folders for site in f['sites']])),
                'date_range': list(set([date for f in folders for date in f['dates']]))
            }
        
        # 사이트별 통계
        site_stats = {}
        for site, folders in self.project_network.items():
            site_stats[site] = {
                'folder_count': len(folders),
                'case_count': len(set([case for f in folders for case in f['cases']])),
                'site_name': self.site_codes.get(site, site)
            }
        
        # 날짜별 통계
        date_stats = {}
        for date, folders in self.date_network.items():
            date_stats[date] = {
                'folder_count': len(folders),
                'case_count': len(set([case for f in folders for case in f['cases']])),
                'phases': list(set([f['phase'] for f in folders]))
            }
        
        # 프로젝트 단계별 통계
        phase_stats = defaultdict(int)
        for mapping in self.folder_mappings:
            phase_stats[mapping['phase']] += 1
        
        # 보고서 생성
        report = f"""
# HVDC 폴더 제목 기반 매핑 분석 보고서
생성일시: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

## 📊 전체 프로젝트 개요
- **총 폴더 수**: {total_folders:,}개
- **총 케이스 수**: {total_cases:,}개
- **총 날짜 수**: {total_dates:,}개
- **총 사이트 수**: {total_sites:,}개

## 🎯 케이스별 상세 분석 (Top 20)
"""
        
        # 케이스별 상위 20개
        if case_stats:
            top_cases = sorted(case_stats.items(), key=lambda x: x[1]['folder_count'], reverse=True)[:20]
        else:
            top_cases = []
        if top_cases:
            for case, stats in top_cases:
                report += f"""
### {case}
- **관련 폴더**: {stats['folder_count']}개
- **프로젝트 단계**: {', '.join(stats['phases'])}
- **관련 사이트**: {', '.join(stats['sites'][:5])}{'...' if len(stats['sites']) > 5 else ''}
- **날짜 범위**: {', '.join(stats['date_range'][:3])}{'...' if len(stats['date_range']) > 3 else ''}
"""
        else:
            report += "\n### 케이스 정보가 없습니다.\n"
        
        report += f"""
## 🏗️ 사이트별 분석
"""
        if site_stats:
            for site, stats in sorted(site_stats.items(), key=lambda x: x[1]['folder_count'], reverse=True):
                report += f"""
### {site} ({stats['site_name']})
- **관련 폴더**: {stats['folder_count']}개
- **케이스 수**: {stats['case_count']}개
"""
        else:
            report += "\n### 사이트 정보가 없습니다.\n"
        
        report += f"""
## 📅 날짜별 분석 (최근 10개)
"""
        if date_stats:
            recent_dates = sorted(date_stats.items(), key=lambda x: x[0], reverse=True)[:10]
            for date, stats in recent_dates:
                report += f"""
### {date}
- **관련 폴더**: {stats['folder_count']}개
- **케이스 수**: {stats['case_count']}개
- **프로젝트 단계**: {', '.join(stats['phases'])}
"""
        else:
            report += "\n### 날짜 정보가 없습니다.\n"
        
        report += f"""
## 📈 프로젝트 단계별 분포
"""
        for phase, count in sorted(phase_stats.items(), key=lambda x: x[1], reverse=True):
            phase_name = {
                'procurement': '조달',
                'shipping': '운송',
                'customs': '관세',
                'logistics': '물류',
                'installation': '설치',
                'testing': '테스트',
                'certification': '인증',
                'general': '일반'
            }.get(phase, phase)
            report += f"- **{phase_name}**: {count:,}개 ({count/total_folders*100:.1f}%)\n"
        
        report += f"""
## 🔗 폴더-케이스-날짜 연결성 분석

### 주요 연결 패턴
1. **폴더 중심 연결**: 각 폴더별로 관련된 케이스, 날짜, 사이트 추적
2. **케이스 중심 연결**: 각 케이스별로 관련된 모든 폴더와 날짜 추적
3. **날짜 중심 연결**: 각 날짜별로 관련된 모든 폴더와 케이스 추적
4. **사이트 중심 연결**: 각 사이트별로 관련된 모든 활동 추적

### 권장사항
1. **데이터 정합성**: 동일 케이스의 폴더들이 일관된 정보를 포함하는지 확인
2. **프로세스 최적화**: 빈번한 폴더 생성이 발생하는 단계의 프로세스 개선
3. **통합 관리**: 분산된 폴더의 관련 정보들을 통합 관리 시스템으로 연결
4. **자동화 구축**: 반복적인 폴더 패턴의 자동 처리 시스템 구축
"""
        
        return report

    def save_mapping_data(self):
        """매핑 데이터 저장"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # 폴더 매핑 데이터 CSV 저장
        df_folders = pd.DataFrame(self.folder_mappings)
        df_folders.to_csv(f'hvdc_folder_mappings_{timestamp}.csv', index=False, encoding='utf-8-sig')
        
        # 케이스별 상세 데이터 저장
        case_details = []
        for case, folders in self.case_network.items():
            case_details.append({
                'case_number': case,
                'folder_count': len(folders),
                'phases': ', '.join(set([f['phase'] for f in folders])),
                'sites': ', '.join(set([site for f in folders for site in f['sites']])),
                'dates': ', '.join(set([date for f in folders for date in f['dates']]))
            })
        
        df_cases = pd.DataFrame(case_details)
        df_cases = df_cases.sort_values('folder_count', ascending=False)
        df_cases.to_csv(f'hvdc_case_folder_analysis_{timestamp}.csv', index=False, encoding='utf-8-sig')
        
        # 사이트별 상세 데이터 저장
        site_details = []
        for site, folders in self.project_network.items():
            site_details.append({
                'site_code': site,
                'site_name': self.site_codes.get(site, site),
                'folder_count': len(folders),
                'case_count': len(set([case for f in folders for case in f['cases']])),
                'phases': ', '.join(set([f['phase'] for f in folders]))
            })
        
        df_sites = pd.DataFrame(site_details)
        df_sites = df_sites.sort_values('folder_count', ascending=False)
        df_sites.to_csv(f'hvdc_site_folder_analysis_{timestamp}.csv', index=False, encoding='utf-8-sig')
        
        # 날짜별 상세 데이터 저장
        date_details = []
        for date, folders in self.date_network.items():
            date_details.append({
                'date': date,
                'folder_count': len(folders),
                'case_count': len(set([case for f in folders for case in f['cases']])),
                'phases': ', '.join(set([f['phase'] for f in folders]))
            })
        
        df_dates = pd.DataFrame(date_details)
        df_dates = df_dates.sort_values('date', ascending=False)
        df_dates.to_csv(f'hvdc_date_folder_analysis_{timestamp}.csv', index=False, encoding='utf-8-sig')
        
        # 종합 보고서 저장
        report = self.generate_comprehensive_report()
        with open(f'hvdc_folder_mapping_report_{timestamp}.md', 'w', encoding='utf-8') as f:
            f.write(report)
        
        logger.info(f"폴더 매핑 데이터 저장 완료: {timestamp}")

def main():
    """메인 실행 함수"""
    email_root = r"C:/Users/SAMSUNG/Documents/EMAIL"
    
    logger.info("HVDC 폴더 제목 기반 매핑 시작")
    
    mapper = FolderTitleMapper(email_root)
    
    try:
        # 모든 폴더 스캔 및 매핑
        mapper.scan_all_folders()
        
        # 매핑 데이터 저장
        mapper.save_mapping_data()
        
        logger.info("HVDC 폴더 제목 기반 매핑 완료")
        
        # 간단한 요약 출력
        print(f"\n🎯 HVDC 폴더 제목 매핑 완료!")
        print(f"📁 총 폴더: {len(mapper.folder_mappings):,}개")
        print(f"🎯 총 케이스: {len(mapper.case_network):,}개")
        print(f"📅 총 날짜: {len(mapper.date_network):,}개")
        print(f"🏗️ 총 사이트: {len(mapper.project_network):,}개")
        
    except Exception as e:
        logger.error(f"폴더 매핑 실행 오류: {str(e)}")
        print(f"❌ 오류 발생: {str(e)}")

if __name__ == "__main__":
    main()
