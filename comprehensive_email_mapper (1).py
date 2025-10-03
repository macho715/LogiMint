#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HVDC 프로젝트 종합 이메일 매핑 시스템
모든 폴더 스캔 후 파일 제목과 날짜별 매핑으로 전체 프로젝트 유기적 연결
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
import networkx as nx
import matplotlib.pyplot as plt
import seaborn as sns

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('comprehensive_mapping.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class ComprehensiveEmailMapper:
    """HVDC 프로젝트 종합 이메일 매핑 클래스"""
    
    def __init__(self, email_root_path: str):
        self.email_root_path = Path(email_root_path)
        self.scan_results = {}
        self.project_connections = defaultdict(list)
        self.timeline_data = []
        self.vendor_network = defaultdict(list)
        self.site_network = defaultdict(list)
        self.case_network = defaultdict(list)
        
        # HVDC 케이스 번호 추출 패턴
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
            'ZEN': 'ZENER',
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
        
        # 프로젝트 단계별 키워드
        self.project_phases = {
            'procurement': ['LPO', 'PO', 'Purchase Order', 'Procurement', 'Order'],
            'shipping': ['Shipping', 'Delivery', 'Container', 'CNTR', 'LCT', 'Vessel'],
            'customs': ['Customs', 'Clearance', 'Import', 'Export', 'Duty'],
            'logistics': ['Logistics', 'Transport', 'Freight', 'Cargo', 'Material'],
            'installation': ['Installation', 'Install', 'Mounting', 'Assembly'],
            'testing': ['Test', 'Testing', 'Commissioning', 'Startup'],
            'certification': ['Certificate', 'Cert', 'MTC', 'COC', 'Quality']
        }

    def load_scan_results(self) -> bool:
        """스캔 결과 로드"""
        try:
            # 최신 스캔 결과 파일 찾기
            result_files = list(Path('.').glob('email_scan_results_*.json'))
            if not result_files:
                logger.error("스캔 결과 파일을 찾을 수 없습니다.")
                return False
            
            latest_file = max(result_files, key=os.path.getctime)
            logger.info(f"스캔 결과 로드: {latest_file}")
            
            with open(latest_file, 'r', encoding='utf-8') as f:
                self.scan_results = json.load(f)
            
            return True
        except Exception as e:
            logger.error(f"스캔 결과 로드 오류: {str(e)}")
            return False

    def extract_case_numbers(self, subject: str) -> List[str]:
        """이메일 제목에서 HVDC 케이스 번호 추출"""
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

    def extract_project_phase(self, subject: str) -> str:
        """이메일 제목에서 프로젝트 단계 추출"""
        subject_lower = str(subject).lower()
        
        for phase, keywords in self.project_phases.items():
            for keyword in keywords:
                if keyword.lower() in subject_lower:
                    return phase
        
        return 'general'

    def extract_lpo_numbers(self, subject: str) -> List[str]:
        """LPO 번호 추출"""
        lpo_pattern = r'LPO[-\s]?(\d+)'
        matches = re.findall(lpo_pattern, str(subject), re.IGNORECASE)
        return [f"LPO-{m}" for m in matches]

    def extract_site_codes(self, subject: str) -> List[str]:
        """사이트 코드 추출"""
        site_pattern = r'\b(DAS|AGI|MIR|MIRFA|GHALLAN)\b'
        matches = re.findall(site_pattern, str(subject), re.IGNORECASE)
        return [match.upper() for match in matches]

    def build_project_connections(self):
        """프로젝트 연결 관계 구축"""
        logger.info("프로젝트 연결 관계 구축 시작")
        
        for folder_name, folder_data in self.scan_results.get('folders', {}).items():
            if 'emails' not in folder_data:
                continue
            
            for email in folder_data['emails']:
                subject = email.get('subject', '')
                received_time = email.get('received_time', '')
                sender = email.get('sender', '')
                
                # 케이스 번호 추출
                cases = self.extract_case_numbers(subject)
                
                # 프로젝트 단계 추출
                phase = self.extract_project_phase(subject)
                
                # LPO 번호 추출
                lpos = self.extract_lpo_numbers(subject)
                
                # 사이트 코드 추출
                sites = self.extract_site_codes(subject)
                
                # 타임라인 데이터 추가
                timeline_entry = {
                    'folder': folder_name,
                    'subject': subject,
                    'sender': sender,
                    'received_time': received_time,
                    'cases': cases,
                    'phase': phase,
                    'lpos': lpos,
                    'sites': sites,
                    'file_path': email.get('file_path', '')
                }
                self.timeline_data.append(timeline_entry)
                
                # 케이스별 연결 관계
                for case in cases:
                    self.case_network[case].append({
                        'folder': folder_name,
                        'subject': subject,
                        'received_time': received_time,
                        'phase': phase
                    })
                
                # 사이트별 연결 관계
                for site in sites:
                    self.site_network[site].append({
                        'folder': folder_name,
                        'subject': subject,
                        'received_time': received_time,
                        'cases': cases
                    })
                
                # 벤더별 연결 관계
                for case in cases:
                    vendor_match = re.search(r'HVDC-ADOPT-([A-Z]+)-', case)
                    if vendor_match:
                        vendor = vendor_match.group(1)
                        self.vendor_network[vendor].append({
                            'folder': folder_name,
                            'subject': subject,
                            'received_time': received_time,
                            'case': case
                        })

    def generate_comprehensive_report(self) -> str:
        """종합 프로젝트 연결 보고서 생성"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # 기본 통계
        total_emails = len(self.timeline_data)
        total_cases = len(self.case_network)
        total_sites = len(self.site_network)
        total_vendors = len(self.vendor_network)
        
        # 케이스별 통계
        case_stats = {}
        for case, emails in self.case_network.items():
            case_stats[case] = {
                'email_count': len(emails),
                'phases': list(set([e['phase'] for e in emails])),
                'folders': list(set([e['folder'] for e in emails])),
                'date_range': {
                    'earliest': min([e['received_time'] for e in emails if e['received_time']]),
                    'latest': max([e['received_time'] for e in emails if e['received_time']])
                }
            }
        
        # 사이트별 통계
        site_stats = {}
        for site, emails in self.site_network.items():
            site_stats[site] = {
                'email_count': len(emails),
                'case_count': len(set([case for e in emails for case in e['cases']])),
                'folders': list(set([e['folder'] for e in emails])),
                'site_name': self.site_codes.get(site, site)
            }
        
        # 벤더별 통계
        vendor_stats = {}
        for vendor, emails in self.vendor_network.items():
            vendor_stats[vendor] = {
                'email_count': len(emails),
                'case_count': len(set([e['case'] for e in emails])),
                'folders': list(set([e['folder'] for e in emails])),
                'vendor_name': self.vendor_codes.get(vendor, vendor)
            }
        
        # 프로젝트 단계별 통계
        phase_stats = defaultdict(int)
        for entry in self.timeline_data:
            phase_stats[entry['phase']] += 1
        
        # 보고서 생성
        report = f"""
# HVDC 프로젝트 종합 이메일 연결 분석 보고서
생성일시: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

## 📊 전체 프로젝트 개요
- **총 이메일 수**: {total_emails:,}개
- **총 케이스 수**: {total_cases:,}개
- **총 사이트 수**: {total_sites:,}개
- **총 벤더 수**: {total_vendors:,}개

## 🎯 케이스별 상세 분석 (Top 20)
"""
        
        # 케이스별 상위 20개
        top_cases = sorted(case_stats.items(), key=lambda x: x[1]['email_count'], reverse=True)[:20]
        for case, stats in top_cases:
            report += f"""
### {case}
- **이메일 수**: {stats['email_count']}개
- **관련 폴더**: {', '.join(stats['folders'][:5])}{'...' if len(stats['folders']) > 5 else ''}
- **프로젝트 단계**: {', '.join(stats['phases'])}
- **기간**: {stats['date_range']['earliest']} ~ {stats['date_range']['latest']}
"""
        
        report += f"""
## 🏗️ 사이트별 분석
"""
        for site, stats in sorted(site_stats.items(), key=lambda x: x[1]['email_count'], reverse=True):
            report += f"""
### {site} ({stats['site_name']})
- **이메일 수**: {stats['email_count']}개
- **케이스 수**: {stats['case_count']}개
- **관련 폴더**: {', '.join(stats['folders'][:5])}{'...' if len(stats['folders']) > 5 else ''}
"""
        
        report += f"""
## 🏢 벤더별 분석
"""
        for vendor, stats in sorted(vendor_stats.items(), key=lambda x: x[1]['email_count'], reverse=True):
            report += f"""
### {vendor} ({stats['vendor_name']})
- **이메일 수**: {stats['email_count']}개
- **케이스 수**: {stats['case_count']}개
- **관련 폴더**: {', '.join(stats['folders'][:5])}{'...' if len(stats['folders']) > 5 else ''}
"""
        
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
            report += f"- **{phase_name}**: {count:,}개 ({count/total_emails*100:.1f}%)\n"
        
        report += f"""
## 🔗 프로젝트 연결성 분석

### 주요 연결 패턴
1. **케이스 중심 연결**: 각 케이스별로 관련된 모든 이메일과 폴더 추적
2. **사이트 중심 연결**: 각 사이트별로 관련된 모든 활동 추적
3. **벤더 중심 연결**: 각 벤더별로 관련된 모든 거래 추적
4. **시간 중심 연결**: 프로젝트 진행 단계별 이메일 흐름 추적

### 권장사항
1. **데이터 정합성**: 동일 케이스의 이메일들이 일관된 정보를 포함하는지 확인
2. **프로세스 최적화**: 빈번한 이메일 교환이 발생하는 단계의 프로세스 개선
3. **통합 관리**: 분산된 폴더의 관련 이메일들을 통합 관리 시스템으로 연결
4. **자동화 구축**: 반복적인 이메일 패턴의 자동 처리 시스템 구축
"""
        
        return report

    def create_network_visualization(self):
        """네트워크 시각화 생성"""
        logger.info("네트워크 시각화 생성 시작")
        
        # NetworkX 그래프 생성
        G = nx.Graph()
        
        # 노드 추가
        for case in self.case_network.keys():
            G.add_node(case, node_type='case')
        
        for site in self.site_network.keys():
            G.add_node(site, node_type='site')
        
        for vendor in self.vendor_network.keys():
            G.add_node(vendor, node_type='vendor')
        
        # 엣지 추가 (케이스-사이트-벤더 연결)
        for case, emails in self.case_network.items():
            for email in emails:
                # 케이스-사이트 연결
                for site in email.get('sites', []):
                    G.add_edge(case, site, weight=1)
                
                # 케이스-벤더 연결
                vendor_match = re.search(r'HVDC-ADOPT-([A-Z]+)-', case)
                if vendor_match:
                    vendor = vendor_match.group(1)
                    G.add_edge(case, vendor, weight=1)
        
        # 시각화
        plt.figure(figsize=(20, 15))
        pos = nx.spring_layout(G, k=3, iterations=50)
        
        # 노드 타입별 색상
        node_colors = []
        for node in G.nodes():
            if G.nodes[node]['node_type'] == 'case':
                node_colors.append('lightblue')
            elif G.nodes[node]['node_type'] == 'site':
                node_colors.append('lightgreen')
            else:  # vendor
                node_colors.append('lightcoral')
        
        # 그래프 그리기
        nx.draw(G, pos, 
                node_color=node_colors,
                node_size=500,
                with_labels=True,
                font_size=8,
                font_weight='bold',
                edge_color='gray',
                alpha=0.7)
        
        plt.title('HVDC 프로젝트 네트워크 연결도', fontsize=16, fontweight='bold')
        plt.axis('off')
        
        # 범례 추가
        legend_elements = [
            plt.Line2D([0], [0], marker='o', color='w', markerfacecolor='lightblue', markersize=10, label='케이스'),
            plt.Line2D([0], [0], marker='o', color='w', markerfacecolor='lightgreen', markersize=10, label='사이트'),
            plt.Line2D([0], [0], marker='o', color='w', markerfacecolor='lightcoral', markersize=10, label='벤더')
        ]
        plt.legend(handles=legend_elements, loc='upper right')
        
        plt.tight_layout()
        plt.savefig('hvdc_project_network.png', dpi=300, bbox_inches='tight')
        plt.close()
        
        logger.info("네트워크 시각화 완료: hvdc_project_network.png")

    def save_comprehensive_data(self):
        """종합 데이터 저장"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # 타임라인 데이터 CSV 저장
        df_timeline = pd.DataFrame(self.timeline_data)
        df_timeline.to_csv(f'hvdc_timeline_data_{timestamp}.csv', index=False, encoding='utf-8-sig')
        
        # 케이스별 상세 데이터 저장
        case_details = []
        for case, emails in self.case_network.items():
            case_details.append({
                'case_number': case,
                'email_count': len(emails),
                'phases': ', '.join(set([e['phase'] for e in emails])),
                'folders': ', '.join(set([e['folder'] for e in emails])),
                'date_range': f"{min([e['received_time'] for e in emails if e['received_time']])} ~ {max([e['received_time'] for e in emails if e['received_time']])}"
            })
        
        df_cases = pd.DataFrame(case_details)
        df_cases = df_cases.sort_values('email_count', ascending=False)
        df_cases.to_csv(f'hvdc_case_analysis_{timestamp}.csv', index=False, encoding='utf-8-sig')
        
        # 사이트별 상세 데이터 저장
        site_details = []
        for site, emails in self.site_network.items():
            site_details.append({
                'site_code': site,
                'site_name': self.site_codes.get(site, site),
                'email_count': len(emails),
                'case_count': len(set([case for e in emails for case in e['cases']])),
                'folders': ', '.join(set([e['folder'] for e in emails]))
            })
        
        df_sites = pd.DataFrame(site_details)
        df_sites = df_sites.sort_values('email_count', ascending=False)
        df_sites.to_csv(f'hvdc_site_analysis_{timestamp}.csv', index=False, encoding='utf-8-sig')
        
        # 벤더별 상세 데이터 저장
        vendor_details = []
        for vendor, emails in self.vendor_network.items():
            vendor_details.append({
                'vendor_code': vendor,
                'vendor_name': self.vendor_codes.get(vendor, vendor),
                'email_count': len(emails),
                'case_count': len(set([e['case'] for e in emails])),
                'folders': ', '.join(set([e['folder'] for e in emails]))
            })
        
        df_vendors = pd.DataFrame(vendor_details)
        df_vendors = df_vendors.sort_values('email_count', ascending=False)
        df_vendors.to_csv(f'hvdc_vendor_analysis_{timestamp}.csv', index=False, encoding='utf-8-sig')
        
        # 종합 보고서 저장
        report = self.generate_comprehensive_report()
        with open(f'hvdc_comprehensive_report_{timestamp}.md', 'w', encoding='utf-8') as f:
            f.write(report)
        
        logger.info(f"종합 데이터 저장 완료: {timestamp}")

def main():
    """메인 실행 함수"""
    email_root = r"C:/Users/SAMSUNG/Documents/EMAIL"
    
    logger.info("HVDC 프로젝트 종합 이메일 매핑 시작")
    
    mapper = ComprehensiveEmailMapper(email_root)
    
    try:
        # 스캔 결과 로드
        if not mapper.load_scan_results():
            logger.error("스캔 결과 로드 실패")
            return
        
        # 프로젝트 연결 관계 구축
        mapper.build_project_connections()
        
        # 네트워크 시각화 생성
        mapper.create_network_visualization()
        
        # 종합 데이터 저장
        mapper.save_comprehensive_data()
        
        logger.info("HVDC 프로젝트 종합 이메일 매핑 완료")
        
        # 간단한 요약 출력
        print(f"\n🎯 HVDC 프로젝트 종합 분석 완료!")
        print(f"📧 총 이메일: {len(mapper.timeline_data):,}개")
        print(f"🎯 총 케이스: {len(mapper.case_network):,}개")
        print(f"🏗️ 총 사이트: {len(mapper.site_network):,}개")
        print(f"🏢 총 벤더: {len(mapper.vendor_network):,}개")
        
    except Exception as e:
        logger.error(f"종합 매핑 실행 오류: {str(e)}")
        print(f"❌ 오류 발생: {str(e)}")

if __name__ == "__main__":
    main()
