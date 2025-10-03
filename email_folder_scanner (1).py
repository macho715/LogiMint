#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HVDC EMAIL 폴더 전체 스캔 및 매핑 시스템
C:/Users/SAMSUNG/Documents/EMAIL 폴더의 모든 하위 폴더를 스캔하여 이메일 데이터를 추출하고 HVDC 온톨로지에 매핑
"""

import os
import re
import json
import pandas as pd
from datetime import datetime
from pathlib import Path
import logging
from typing import Dict, List, Tuple, Optional
# Outlook 관련 import 제거 (Outlook 파일 스캔하지 않음)
# import win32com.client
# import pythoncom

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('email_scan.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class EmailFolderScanner:
    """EMAIL 폴더 전체 스캔 및 이메일 추출 클래스"""
    
    def __init__(self, email_root_path: str):
        self.email_root_path = Path(email_root_path)
        self.scan_results = {}
        self.email_data = []
        self.folder_stats = {}
        
        # HVDC 케이스 번호 추출 패턴 (기존 로직 활용)
        self.case_patterns = {
            'pattern1': r'HVDC-ADOPT-([A-Z]+)-([A-Z0-9\-]+)',
            'pattern2': r'HVDC-([A-Z]+)-([A-Z]+)-([A-Z0-9\-]+)',
            'pattern3_outer': r'\(([^\)]+)\)',
            'pattern3_inner': r'([A-Z]+)-([0-9]+(?:-[0-9A-Z]+)?)',
            'pattern4': r'\[HVDC-AGI\].*?(JPTW-(\d+))\s*/\s*(GRM-(\d+))',
            'pattern5': r':\s*([A-Z]+-[A-Z]+-[A-Z]+\d+-[A-Z]+\d+)'
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
            'DAS': 'DAS Site',
            'AGI': 'AGI Site',
            'MIR': 'MIR Site',
            'MIRFA': 'MIRFA Site',
            'GHALLAN': 'GHALLAN Site'
        }
        
        # 사이트 코드 매핑
        self.site_codes = {
            'DAS': '변전소',
            'AGI': '골재 저장소', 
            'MIR': '중간 변전소',
            'MIRFA': '보조 변전소',
            'GHALLAN': '사막 사이트'
        }

    def scan_all_folders(self) -> Dict:
        """EMAIL 폴더의 모든 하위 폴더를 스캔"""
        logger.info(f"EMAIL 폴더 스캔 시작: {self.email_root_path}")
        
        if not self.email_root_path.exists():
            logger.error(f"EMAIL 폴더가 존재하지 않습니다: {self.email_root_path}")
            return {}
        
        folder_count = 0
        total_emails = 0
        
        for folder_path in self.email_root_path.iterdir():
            if folder_path.is_dir():
                folder_count += 1
                logger.info(f"폴더 스캔 중: {folder_path.name}")
                
                folder_result = self.scan_folder(folder_path)
                self.scan_results[folder_path.name] = folder_result
                
                if 'email_count' in folder_result:
                    total_emails += folder_result['email_count']
                
                # 폴더별 통계 저장
                self.folder_stats[folder_path.name] = {
                    'path': str(folder_path),
                    'email_count': folder_result.get('email_count', 0),
                    'case_count': folder_result.get('case_count', 0),
                    'last_modified': folder_path.stat().st_mtime
                }
        
        logger.info(f"스캔 완료: {folder_count}개 폴더, {total_emails}개 이메일")
        
        return {
            'total_folders': folder_count,
            'total_emails': total_emails,
            'scan_timestamp': datetime.now().isoformat(),
            'folders': self.scan_results
        }

    def scan_folder(self, folder_path: Path) -> Dict:
        """개별 폴더 스캔"""
        try:
            # Outlook 폴더인지 확인
            if self.is_outlook_folder(folder_path):
                return self.scan_outlook_folder(folder_path)
            else:
                return self.scan_file_system_folder(folder_path)
        except Exception as e:
            logger.error(f"폴더 스캔 오류 {folder_path.name}: {str(e)}")
            return {'error': str(e), 'email_count': 0, 'case_count': 0}

    def is_outlook_folder(self, folder_path: Path) -> bool:
        """Outlook 폴더인지 확인 - Outlook 파일은 스캔하지 않음"""
        # Outlook 파일은 스캔하지 않도록 항상 False 반환
        return False

    def scan_outlook_folder(self, folder_path: Path) -> Dict:
        """Outlook 폴더 스캔 - 사용하지 않음 (Outlook 파일 제외)"""
        logger.info(f"Outlook 폴더 스킵: {folder_path.name}")
        return {
            'type': 'outlook_skipped',
            'email_count': 0,
            'case_count': 0,
            'emails': []
        }

    def scan_file_system_folder(self, folder_path: Path) -> Dict:
        """파일 시스템 폴더 스캔 - Outlook 파일 제외"""
        email_count = 0
        case_count = 0
        emails = []
        
        # 이메일 관련 파일 확장자 (Outlook 파일 제외)
        email_extensions = ['.eml', '.txt', '.html', '.htm', '.csv']
        
        # Outlook 파일 확장자 제외
        outlook_extensions = ['.pst', '.ost', '.msg']
        
        for file_path in folder_path.rglob('*'):
            if file_path.is_file():
                file_ext = file_path.suffix.lower()
                
                # Outlook 파일은 스킵
                if file_ext in outlook_extensions:
                    logger.info(f"Outlook 파일 스킵: {file_path}")
                    continue
                
                # 이메일 관련 파일만 처리
                if file_ext in email_extensions:
                    try:
                        email_data = self.read_email_file(file_path)
                        if email_data:
                            emails.append(email_data)
                            email_count += 1
                            
                            # 케이스 번호 추출
                            cases = self.extract_case_numbers(email_data.get('subject', ''))
                            if cases:
                                case_count += len(cases)
                                
                    except Exception as e:
                        logger.warning(f"파일 읽기 오류 {file_path}: {str(e)}")
        
        return {
            'type': 'filesystem',
            'email_count': email_count,
            'case_count': case_count,
            'emails': emails[:100]  # 샘플만 저장
        }

    def read_email_file(self, file_path: Path) -> Optional[Dict]:
        """이메일 파일 읽기"""
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
            
            # 기본 이메일 정보 추출
            email_data = {
                'file_path': str(file_path),
                'file_name': file_path.name,
                'file_size': file_path.stat().st_size,
                'modified_time': datetime.fromtimestamp(file_path.stat().st_mtime).isoformat()
            }
            
            # 제목 추출 (간단한 패턴)
            subject_match = re.search(r'Subject:\s*(.+)', content, re.IGNORECASE)
            if subject_match:
                email_data['subject'] = subject_match.group(1).strip()
            
            # 발신자 추출
            from_match = re.search(r'From:\s*(.+)', content, re.IGNORECASE)
            if from_match:
                email_data['sender'] = from_match.group(1).strip()
            
            # 수신 시간 추출
            date_match = re.search(r'Date:\s*(.+)', content, re.IGNORECASE)
            if date_match:
                email_data['received_time'] = date_match.group(1).strip()
            
            return email_data
            
        except Exception as e:
            logger.warning(f"이메일 파일 읽기 실패 {file_path}: {str(e)}")
            return None

    def extract_email_data(self, outlook_item) -> Optional[Dict]:
        """Outlook 아이템에서 이메일 데이터 추출 - 사용하지 않음"""
        return None

    def extract_case_numbers(self, subject: str) -> List[str]:
        """이메일 제목에서 HVDC 케이스 번호 추출 (기존 로직 활용)"""
        if not subject:
            return []
        
        case_numbers = []
        subject_str = str(subject).upper()
        
        # 패턴 1: 완전한 HVDC-ADOPT 패턴
        pattern1 = re.findall(self.case_patterns['pattern1'], subject_str, re.IGNORECASE)
        for match in pattern1:
            vendor_code = match[0].upper()
            case_num = match[1]
            full_case = f"HVDC-ADOPT-{vendor_code}-{case_num}"
            if full_case not in case_numbers:
                case_numbers.append(full_case)
        
        # 패턴 2: HVDC 프로젝트 코드 패턴
        pattern2 = re.findall(self.case_patterns['pattern2'], subject_str, re.IGNORECASE)
        for match in pattern2:
            site_code = match[0].upper()
            vendor_code = match[1].upper()
            case_num = match[2]
            full_case = f"HVDC-{site_code}-{vendor_code}-{case_num}"
            if full_case not in case_numbers:
                case_numbers.append(full_case)
        
        # 패턴 3: 괄호 안 약식 패턴
        pattern3_outer = re.findall(self.case_patterns['pattern3_outer'], subject_str)
        for outer_match in pattern3_outer:
            pattern3_inner = re.findall(self.case_patterns['pattern3_inner'], outer_match, re.IGNORECASE)
            for match in pattern3_inner:
                vendor_code = match[0].upper()
                case_num = match[1]
                full_case = f"HVDC-ADOPT-{vendor_code}-{case_num}"
                if full_case not in case_numbers:
                    case_numbers.append(full_case)
        
        # 패턴 4: JPTW/GRM 패턴
        pattern4 = re.findall(self.case_patterns['pattern4'], subject_str, re.IGNORECASE)
        for match in pattern4:
            jptw_num = match[1]
            grm_num = match[3]
            full_case = f"HVDC-AGI-JPTW{jptw_num}-GRM{grm_num}".upper()
            if full_case not in case_numbers:
                case_numbers.append(full_case)
        
        # 패턴 5: 콜론 뒤 완성형 패턴
        pattern5 = re.findall(self.case_patterns['pattern5'], subject_str, re.IGNORECASE)
        for match in pattern5:
            clean_case = re.sub(r'\(.*?\)', '', match).strip().upper()
            if clean_case not in case_numbers:
                case_numbers.append(clean_case)
        
        return case_numbers

    def generate_comprehensive_report(self) -> str:
        """종합 스캔 보고서 생성"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # 폴더별 통계 생성
        folder_summary = []
        total_emails = 0
        total_cases = 0
        
        for folder_name, stats in self.folder_stats.items():
            folder_summary.append({
                '폴더명': folder_name,
                '이메일수': stats['email_count'],
                '케이스수': stats['case_count'],
                '마지막수정': datetime.fromtimestamp(stats['last_modified']).strftime("%Y-%m-%d %H:%M")
            })
            total_emails += stats['email_count']
            total_cases += stats['case_count']
        
        # 벤더별 통계
        vendor_stats = {}
        site_stats = {}
        
        for folder_name, folder_data in self.scan_results.items():
            if 'emails' in folder_data:
                for email in folder_data['emails']:
                    cases = self.extract_case_numbers(email.get('subject', ''))
                    for case in cases:
                        # 벤더 코드 추출
                        vendor_match = re.search(r'HVDC-ADOPT-([A-Z]+)-', case)
                        if vendor_match:
                            vendor = vendor_match.group(1)
                            vendor_stats[vendor] = vendor_stats.get(vendor, 0) + 1
                        
                        # 사이트 코드 추출
                        site_match = re.search(r'HVDC-([A-Z]+)-', case)
                        if site_match:
                            site = site_match.group(1)
                            site_stats[site] = site_stats.get(site, 0) + 1
        
        # 보고서 생성
        report = f"""
# HVDC EMAIL 폴더 전체 스캔 보고서
생성일시: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

## 📊 전체 통계
- **총 폴더 수**: {len(self.folder_stats)}개
- **총 이메일 수**: {total_emails:,}개
- **총 케이스 수**: {total_cases:,}개
- **케이스 추출률**: {(total_cases/total_emails*100):.1f}% (이메일 대비)

## 📁 폴더별 상세 통계
"""
        
        # 폴더별 통계 테이블
        df_folders = pd.DataFrame(folder_summary)
        df_folders = df_folders.sort_values('이메일수', ascending=False)
        report += df_folders.to_string(index=False)
        
        # 벤더별 통계
        if vendor_stats:
            report += f"\n\n## 🏢 벤더별 케이스 분포\n"
            for vendor, count in sorted(vendor_stats.items(), key=lambda x: x[1], reverse=True):
                vendor_name = self.vendor_codes.get(vendor, vendor)
                report += f"- **{vendor}** ({vendor_name}): {count:,}개\n"
        
        # 사이트별 통계
        if site_stats:
            report += f"\n\n## 🏗️ 사이트별 케이스 분포\n"
            for site, count in sorted(site_stats.items(), key=lambda x: x[1], reverse=True):
                site_name = self.site_codes.get(site, site)
                report += f"- **{site}** ({site_name}): {count:,}개\n"
        
        # 상위 폴더 분석
        report += f"\n\n## 🔝 상위 이메일 폴더 (Top 10)\n"
        top_folders = df_folders.head(10)
        for _, row in top_folders.iterrows():
            report += f"- **{row['폴더명']}**: {row['이메일수']:,}개 이메일, {row['케이스수']:,}개 케이스\n"
        
        # 권장사항
        report += f"""
## 💡 권장사항

### 1. 데이터 정리
- 이메일이 많은 폴더: {top_folders.iloc[0]['폴더명']} ({top_folders.iloc[0]['이메일수']:,}개)
- 케이스 추출률이 낮은 폴더 확인 필요

### 2. 매핑 최적화
- 벤더별 패턴 분석으로 추출률 향상
- 사이트별 특수 패턴 추가 고려

### 3. 자동화 구축
- 정기적 스캔 스케줄 설정
- 실시간 케이스 추출 시스템 구축
"""
        
        return report

    def save_results(self):
        """스캔 결과 저장"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # JSON 결과 저장
        with open(f'email_scan_results_{timestamp}.json', 'w', encoding='utf-8') as f:
            json.dump(self.scan_results, f, ensure_ascii=False, indent=2)
        
        # 폴더 통계 CSV 저장
        df_folders = pd.DataFrame([
            {
                '폴더명': name,
                '이메일수': stats['email_count'],
                '케이스수': stats['case_count'],
                '마지막수정': datetime.fromtimestamp(stats['last_modified']).strftime("%Y-%m-%d %H:%M"),
                '경로': stats['path']
            }
            for name, stats in self.folder_stats.items()
        ])
        df_folders.to_csv(f'email_folder_stats_{timestamp}.csv', index=False, encoding='utf-8-sig')
        
        # 종합 보고서 저장
        report = self.generate_comprehensive_report()
        with open(f'email_scan_report_{timestamp}.md', 'w', encoding='utf-8') as f:
            f.write(report)
        
        logger.info(f"스캔 결과 저장 완료: {timestamp}")

def main():
    """메인 실행 함수"""
    email_root = r"C:/Users/SAMSUNG/Documents/EMAIL"
    
    logger.info("HVDC EMAIL 폴더 전체 스캔 시작")
    
    scanner = EmailFolderScanner(email_root)
    
    try:
        # 전체 폴더 스캔
        scan_results = scanner.scan_all_folders()
        
        # 결과 저장
        scanner.save_results()
        
        logger.info("EMAIL 폴더 스캔 완료")
        
        # 간단한 요약 출력
        print(f"\n📧 EMAIL 폴더 스캔 완료!")
        print(f"📁 총 폴더: {scan_results.get('total_folders', 0)}개")
        print(f"📨 총 이메일: {scan_results.get('total_emails', 0):,}개")
        print(f"🎯 총 케이스: {sum(stats.get('case_count', 0) for stats in scanner.folder_stats.values()):,}개")
        
    except Exception as e:
        logger.error(f"스캔 실행 오류: {str(e)}")
        print(f"❌ 오류 발생: {str(e)}")

if __name__ == "__main__":
    main()
