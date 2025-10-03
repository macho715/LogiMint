#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HVDC 폴더 제목 간단 분석기
폴더 제목에서 케이스 번호, 날짜, 프로젝트 정보를 추출하여 매핑
"""

import os
import re
import json
import pandas as pd
from datetime import datetime
from pathlib import Path
import logging
from typing import Dict, List, Tuple, Optional
from collections import defaultdict, Counter

# 로깅 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class SimpleFolderAnalyzer:
    """간단한 폴더 분석 클래스"""
    
    def __init__(self, email_root_path: str):
        self.email_root_path = Path(email_root_path)
        self.folder_data = []
        
        # HVDC 케이스 번호 추출 패턴
        self.case_patterns = [
            r'HVDC-ADOPT-([A-Z]+)-([A-Z0-9\-]+)',
            r'HVDC-([A-Z]+)-([A-Z]+)-([A-Z0-9\-]+)',
            r'\(([A-Z]+)-(\d+)\)',  # (HE-XXXX)
            r'PRL-([A-Z]+)-(\d+)-([A-Z]+)-([A-Z0-9\-]+)',
            r'\[HVDC-AGI\].*?(JPTW-(\d+))\s*/\s*(GRM-(\d+))'
        ]
        
        # 날짜 패턴
        self.date_patterns = [
            r'(\d{4}-\d{2}-\d{2})',  # YYYY-MM-DD
            r'(\d{4}\.\d{2}\.\d{2})',  # YYYY.MM.DD
            r'(\d{2}-\d{2}-\d{4})',  # MM-DD-YYYY
            r'(\d{4}년\s*\d{1,2}월\s*\d{1,2}일)',  # YYYY년 MM월 DD일
            r'(\d{4}-\d{2}-\d{2}\s*오[전후]\s*\d{1,2}_\d{2}_\d{2})'  # YYYY-MM-DD 오전/오후 HH_MM_SS
        ]

    def extract_cases_from_folder(self, folder_name: str) -> List[str]:
        """폴더명에서 케이스 번호 추출"""
        cases = []
        folder_upper = folder_name.upper()
        
        for pattern in self.case_patterns:
            matches = re.findall(pattern, folder_upper)
            for match in matches:
                if len(match) == 2:  # HVDC-ADOPT-XX-XXXX or (XX-XXXX)
                    vendor, case_num = match
                    cases.append(f"HVDC-ADOPT-{vendor}-{case_num}")
                elif len(match) == 3:  # HVDC-XX-XX-XXXX
                    site, vendor, case_num = match
                    cases.append(f"HVDC-{site}-{vendor}-{case_num}")
                elif len(match) == 4:  # PRL-XX-XX-XX-XXXX or JPTW/GRM
                    if 'JPTW' in str(match):
                        jptw_num, grm_num = match[1], match[3]
                        cases.append(f"HVDC-AGI-JPTW{jptw_num}-GRM{grm_num}")
                    else:
                        site, prl_num, prl_type, prl_desc = match
                        cases.append(f"PRL-{site}-{prl_num}-{prl_type}-{prl_desc}")
        
        return list(set(cases))  # 중복 제거

    def extract_dates_from_folder(self, folder_name: str) -> List[str]:
        """폴더명에서 날짜 추출"""
        dates = []
        
        for pattern in self.date_patterns:
            matches = re.findall(pattern, folder_name)
            dates.extend(matches)
        
        return list(set(dates))  # 중복 제거

    def extract_sites_from_folder(self, folder_name: str) -> List[str]:
        """폴더명에서 사이트 코드 추출"""
        site_pattern = r'\b(DAS|AGI|MIR|MIRFA|GHALLAN|SHU)\b'
        matches = re.findall(site_pattern, folder_name.upper())
        return list(set(matches))

    def extract_lpos_from_folder(self, folder_name: str) -> List[str]:
        """폴더명에서 LPO 번호 추출"""
        lpo_pattern = r'LPO[-\s]?(\d+)'
        matches = re.findall(lpo_pattern, folder_name.upper())
        return [f"LPO-{m}" for m in matches]

    def analyze_folders(self):
        """모든 폴더 분석"""
        logger.info(f"폴더 분석 시작: {self.email_root_path}")
        
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
                
                # 각종 정보 추출
                cases = self.extract_cases_from_folder(folder_name)
                dates = self.extract_dates_from_folder(folder_name)
                sites = self.extract_sites_from_folder(folder_name)
                lpos = self.extract_lpos_from_folder(folder_name)
                
                # 폴더 데이터 저장
                folder_info = {
                    'folder_name': folder_name,
                    'folder_path': str(folder_path),
                    'cases': cases,
                    'dates': dates,
                    'sites': sites,
                    'lpos': lpos,
                    'case_count': len(cases),
                    'date_count': len(dates),
                    'site_count': len(sites),
                    'lpo_count': len(lpos),
                    'last_modified': datetime.fromtimestamp(folder_path.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
                }
                
                self.folder_data.append(folder_info)
                total_cases += len(cases)
                total_dates += len(dates)
                
                if cases or dates or sites or lpos:
                    logger.info(f"📁 {folder_name}: {len(cases)}케이스, {len(dates)}날짜, {len(sites)}사이트, {len(lpos)}LPO")
        
        logger.info(f"폴더 분석 완료: {folder_count}개 폴더, {total_cases}개 케이스, {total_dates}개 날짜")

    def generate_summary_report(self) -> str:
        """요약 보고서 생성"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # 기본 통계
        total_folders = len(self.folder_data)
        folders_with_cases = len([f for f in self.folder_data if f['case_count'] > 0])
        folders_with_dates = len([f for f in self.folder_data if f['date_count'] > 0])
        folders_with_sites = len([f for f in self.folder_data if f['site_count'] > 0])
        folders_with_lpos = len([f for f in self.folder_data if f['lpo_count'] > 0])
        
        # 케이스별 통계
        all_cases = []
        for folder in self.folder_data:
            all_cases.extend(folder['cases'])
        case_counter = Counter(all_cases)
        
        # 사이트별 통계
        all_sites = []
        for folder in self.folder_data:
            all_sites.extend(folder['sites'])
        site_counter = Counter(all_sites)
        
        # 날짜별 통계
        all_dates = []
        for folder in self.folder_data:
            all_dates.extend(folder['dates'])
        date_counter = Counter(all_dates)
        
        # LPO별 통계
        all_lpos = []
        for folder in self.folder_data:
            all_lpos.extend(folder['lpos'])
        lpo_counter = Counter(all_lpos)
        
        # 보고서 생성
        report = f"""
# HVDC 폴더 제목 분석 보고서
생성일시: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

## 📊 전체 통계
- **총 폴더 수**: {total_folders:,}개
- **케이스 포함 폴더**: {folders_with_cases:,}개 ({folders_with_cases/total_folders*100:.1f}%)
- **날짜 포함 폴더**: {folders_with_dates:,}개 ({folders_with_dates/total_folders*100:.1f}%)
- **사이트 포함 폴더**: {folders_with_sites:,}개 ({folders_with_sites/total_folders*100:.1f}%)
- **LPO 포함 폴더**: {folders_with_lpos:,}개 ({folders_with_lpos/total_folders*100:.1f}%)

## 🎯 케이스별 분석 (Top 10)
"""
        
        if case_counter:
            for case, count in case_counter.most_common(10):
                report += f"- **{case}**: {count}개 폴더\n"
        else:
            report += "- 케이스 정보가 없습니다.\n"
        
        report += f"""
## 🏗️ 사이트별 분석
"""
        if site_counter:
            for site, count in site_counter.most_common():
                report += f"- **{site}**: {count}개 폴더\n"
        else:
            report += "- 사이트 정보가 없습니다.\n"
        
        report += f"""
## 📅 날짜별 분석 (최근 10개)
"""
        if date_counter:
            recent_dates = sorted(date_counter.items(), key=lambda x: x[0], reverse=True)[:10]
            for date, count in recent_dates:
                report += f"- **{date}**: {count}개 폴더\n"
        else:
            report += "- 날짜 정보가 없습니다.\n"
        
        report += f"""
## 📋 LPO별 분석 (Top 10)
"""
        if lpo_counter:
            for lpo, count in lpo_counter.most_common(10):
                report += f"- **{lpo}**: {count}개 폴더\n"
        else:
            report += "- LPO 정보가 없습니다.\n"
        
        report += f"""
## 📁 상세 폴더 정보
"""
        for folder in self.folder_data:
            if folder['case_count'] > 0 or folder['date_count'] > 0 or folder['site_count'] > 0:
                report += f"""
### {folder['folder_name']}
- **케이스**: {', '.join(folder['cases'][:5])}{'...' if len(folder['cases']) > 5 else ''}
- **날짜**: {', '.join(folder['dates'][:3])}{'...' if len(folder['dates']) > 3 else ''}
- **사이트**: {', '.join(folder['sites'])}
- **LPO**: {', '.join(folder['lpos'][:3])}{'...' if len(folder['lpos']) > 3 else ''}
"""
        
        return report

    def save_results(self):
        """결과 저장"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # CSV 저장
        df = pd.DataFrame(self.folder_data)
        df.to_csv(f'hvdc_folder_analysis_{timestamp}.csv', index=False, encoding='utf-8-sig')
        
        # JSON 저장
        with open(f'hvdc_folder_analysis_{timestamp}.json', 'w', encoding='utf-8') as f:
            json.dump(self.folder_data, f, ensure_ascii=False, indent=2)
        
        # 보고서 저장
        report = self.generate_summary_report()
        with open(f'hvdc_folder_analysis_report_{timestamp}.md', 'w', encoding='utf-8') as f:
            f.write(report)
        
        logger.info(f"결과 저장 완료: {timestamp}")

def main():
    """메인 실행 함수"""
    email_root = r"C:/Users/SAMSUNG/Documents/EMAIL"
    
    logger.info("HVDC 폴더 제목 분석 시작")
    
    analyzer = SimpleFolderAnalyzer(email_root)
    
    try:
        # 폴더 분석
        analyzer.analyze_folders()
        
        # 결과 저장
        analyzer.save_results()
        
        logger.info("HVDC 폴더 제목 분석 완료")
        
        # 간단한 요약 출력
        print(f"\n🎯 HVDC 폴더 제목 분석 완료!")
        print(f"📁 총 폴더: {len(analyzer.folder_data):,}개")
        
        # 케이스가 있는 폴더들
        folders_with_cases = [f for f in analyzer.folder_data if f['case_count'] > 0]
        print(f"🎯 케이스 포함 폴더: {len(folders_with_cases):,}개")
        
        # 날짜가 있는 폴더들
        folders_with_dates = [f for f in analyzer.folder_data if f['date_count'] > 0]
        print(f"📅 날짜 포함 폴더: {len(folders_with_dates):,}개")
        
        # 사이트가 있는 폴더들
        folders_with_sites = [f for f in analyzer.folder_data if f['site_count'] > 0]
        print(f"🏗️ 사이트 포함 폴더: {len(folders_with_sites):,}개")
        
    except Exception as e:
        logger.error(f"폴더 분석 실행 오류: {str(e)}")
        print(f"❌ 오류 발생: {str(e)}")

if __name__ == "__main__":
    main()
