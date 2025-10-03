#!/usr/bin/env python3
"""
HVDC-ADOPT 화물 추적 시스템
MACHO-GPT v3.4-mini - Cargo Tracking Intelligence

기능:
- HVDC-ADOPT-XXX-XXXX 케이스별 ETA/ATA 추적
- 현장 입고 날짜 (SHU/MIR/DAS/AGI) 모니터링
- 화물 상태 (STATUS) 자동 분류
- 이메일 연동 추적
- 온톨로지 매핑 통합
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from pathlib import Path
import json
import re
from typing import Dict, List, Tuple


class HVDCCargoTracker:
    """HVDC-ADOPT 화물 추적 시스템"""
    
    def __init__(self, 
                 status_file: str = "HVDC PJT/hvdc_macho_gpt/HVDC STATUS/data/HVDC STATUS(20250810).xlsx",
                 email_json: str = "hvdc_emails_structured.json"):
        self.status_file = status_file
        self.email_json = email_json
        self.df_status = None
        self.df_emails = None
        self.tracking_data = []
        
        # 사이트 컬럼
        self.site_columns = ['SHU', 'MIR', 'DAS', 'AGI']
        
        # 창고 컬럼
        self.warehouse_columns = [
            'DSV\nIndoor', 'DSV\nOutdoor', 'DSV\nMZD', 
            'JDN\nMZD', 'JDN\nWaterfront', 'MOSB', 
            'AAA Storage', 'ZENER (WH)', 'Hauler DG Storage', 'Vijay Tanks'
        ]
        
        # 상태 정의
        self.status_map = {
            'NOT_SHIPPED': '미출고',
            'IN_TRANSIT': '운송중',
            'PORT_ARRIVAL': '항구도착',
            'CUSTOMS': '통관중',
            'WAREHOUSE': '창고보관',
            'SITE_DELIVERED': '현장인도',
            'DELAYED': '지연',
            'UNKNOWN': '상태불명'
        }
    
    def load_data(self):
        """데이터 로드"""
        print("="*70)
        print("📦 HVDC-ADOPT 화물 추적 시스템 초기화")
        print("="*70)
        
        # STATUS 파일 로드
        print(f"\n📂 STATUS 파일 로드: {self.status_file}")
        try:
            self.df_status = pd.read_excel(self.status_file)
            print(f"   ✅ {len(self.df_status):,}건의 화물 데이터 로드")
            
            # 날짜 컬럼 변환
            date_columns = ['ETD', 'ATD', 'ETA', 'ATA '] + self.site_columns + self.warehouse_columns
            for col in date_columns:
                if col in self.df_status.columns:
                    self.df_status[col] = pd.to_datetime(self.df_status[col], errors='coerce')
            
            print(f"   📅 날짜 컬럼 변환 완료")
        except Exception as e:
            print(f"   ❌ 오류: {e}")
            return False
        
        # 이메일 데이터 로드
        if Path(self.email_json).exists():
            print(f"\n📧 이메일 데이터 로드: {self.email_json}")
            try:
                with open(self.email_json, 'r', encoding='utf-8') as f:
                    email_data = json.load(f)
                    self.df_emails = pd.DataFrame(email_data['emails'])
                    print(f"   ✅ {len(self.df_emails):,}건의 이메일 데이터 로드")
            except Exception as e:
                print(f"   ⚠️  이메일 데이터 로드 실패: {e}")
        
        return True
    
    def determine_status(self, row: pd.Series) -> str:
        """화물 상태 결정"""
        # 1. 현장 인도 확인
        for site in self.site_columns:
            if pd.notna(row.get(site)):
                return 'SITE_DELIVERED'
        
        # 2. 창고 보관 확인
        for warehouse in self.warehouse_columns:
            if pd.notna(row.get(warehouse)):
                return 'WAREHOUSE'
        
        # 3. 통관 확인
        if pd.notna(row.get('Customs\nStart')):
            if pd.isna(row.get('Customs\nClose')):
                return 'CUSTOMS'
        
        # 4. 항구 도착 확인
        if pd.notna(row.get('ATA ')):
            return 'PORT_ARRIVAL'
        
        # 5. 운송 중 확인
        if pd.notna(row.get('ATD')) or pd.notna(row.get('ETD')):
            # 지연 체크
            eta = row.get('ETA')
            if pd.notna(eta) and eta < pd.Timestamp.now():
                if pd.isna(row.get('ATA ')):
                    return 'DELAYED'
            return 'IN_TRANSIT'
        
        # 6. 미출고
        if pd.isna(row.get('ETD')) and pd.isna(row.get('ATD')):
            return 'NOT_SHIPPED'
        
        return 'UNKNOWN'
    
    def calculate_lead_time(self, row: pd.Series) -> Dict:
        """리드타임 계산"""
        lead_times = {}
        
        # ATD -> ATA
        if pd.notna(row.get('ATD')) and pd.notna(row.get('ATA ')):
            lead_times['shipping_days'] = (row['ATA '] - row['ATD']).days
        
        # ATA -> 첫 번째 사이트
        ata = row.get('ATA ')
        if pd.notna(ata):
            for site in self.site_columns:
                site_date = row.get(site)
                if pd.notna(site_date):
                    lead_times[f'port_to_{site.lower()}'] = (site_date - ata).days
                    break
        
        # 총 리드타임 (ATD -> 최종 사이트)
        atd = row.get('ATD')
        if pd.notna(atd):
            site_dates = [row.get(s) for s in self.site_columns if pd.notna(row.get(s))]
            if site_dates:
                final_site_date = max(site_dates)
                lead_times['total_days'] = (final_site_date - atd).days
        
        return lead_times
    
    def get_current_location(self, row: pd.Series) -> str:
        """현재 위치 결정"""
        # 최근 날짜 기준으로 위치 결정
        location_dates = {}
        
        # 사이트
        for site in self.site_columns:
            if pd.notna(row.get(site)):
                location_dates[site] = row[site]
        
        # 창고
        for warehouse in self.warehouse_columns:
            if pd.notna(row.get(warehouse)):
                warehouse_name = warehouse.replace('\n', ' ').strip()
                location_dates[warehouse_name] = row[warehouse]
        
        if location_dates:
            latest_location = max(location_dates.items(), key=lambda x: x[1])
            return latest_location[0]
        
        # 위치 없으면 상태 기반
        if pd.notna(row.get('ATA ')):
            return 'PORT'
        elif pd.notna(row.get('ATD')):
            return 'SEA/AIR'
        else:
            return 'ORIGIN'
    
    def find_related_emails(self, case_no: str) -> List[Dict]:
        """관련 이메일 찾기"""
        if self.df_emails is None:
            return []
        
        related = []
        # 케이스 번호로 필터링
        case_pattern = case_no.replace('HVDC-ADOPT-', '')
        
        for idx, email in self.df_emails.iterrows():
            subject = str(email.get('subject', '')).upper()
            if case_pattern in subject or case_no in subject:
                related.append({
                    'date': email.get('received_time'),
                    'sender': email.get('sender_name'),
                    'subject': email.get('subject'),
                    'type': email.get('email_type')
                })
        
        return related
    
    def process_tracking_data(self):
        """화물 추적 데이터 처리"""
        print("\n🔄 화물 추적 데이터 처리 중...")
        
        for idx, row in self.df_status.iterrows():
            if idx % 100 == 0:
                print(f"   진행률: {idx}/{len(self.df_status)} ({idx/len(self.df_status)*100:.1f}%)")
            
            case_no = row.get('HVDC CODE')
            if pd.isna(case_no) or not str(case_no).startswith('HVDC-ADOPT-'):
                continue
            
            # 상태 결정
            status = self.determine_status(row)
            
            # 리드타임 계산
            lead_times = self.calculate_lead_time(row)
            
            # 현재 위치
            current_location = self.get_current_location(row)
            
            # 관련 이메일
            related_emails = self.find_related_emails(case_no)
            
            # 추적 데이터 생성
            tracking_info = {
                'case_no': case_no,
                'status': status,
                'status_kr': self.status_map[status],
                'current_location': current_location,
                
                # 기본 정보
                'vendor': row.get('VENDOR'),
                'description': row.get('SUB DESCRIPTION'),
                'bl_no': row.get('B/L No./\nAWB No.'),
                'vessel': row.get('VESSEL NAME/\nFLIGHT No.'),
                
                # 날짜 정보
                'etd': row.get('ETD'),
                'atd': row.get('ATD'),
                'eta': row.get('ETA'),
                'ata': row.get('ATA '),
                
                # 현장 입고 날짜
                'site_shu': row.get('SHU'),
                'site_mir': row.get('MIR'),
                'site_das': row.get('DAS'),
                'site_agi': row.get('AGI'),
                
                # 리드타임
                'lead_times': lead_times,
                
                # 패키지 정보
                'pkg': row.get('PKG'),
                'cbm': row.get('CBM'),
                'gwt': row.get('GWT\n(KG)'),
                
                # 관련 이메일
                'related_emails_count': len(related_emails),
                'latest_email': related_emails[0] if related_emails else None
            }
            
            self.tracking_data.append(tracking_info)
        
        print(f"   ✅ {len(self.tracking_data):,}건의 HVDC-ADOPT 화물 처리 완료")
    
    def generate_summary_report(self, output_file: str = "hvdc_cargo_tracking_report.md"):
        """요약 리포트 생성"""
        print(f"\n📊 추적 리포트 생성: {output_file}")
        
        df = pd.DataFrame(self.tracking_data)
        
        md = []
        md.append("# HVDC-ADOPT 화물 추적 리포트")
        md.append("\n**MACHO-GPT v3.4-mini** | Samsung C&T Logistics × ADNOC·DSV")
        md.append(f"\n**생성일:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        md.append(f"\n**총 화물:** {len(df):,}건")
        md.append(f"\n---\n")
        
        # 상태별 통계
        md.append("## 📊 화물 상태 분포")
        status_counts = df['status_kr'].value_counts()
        for status, count in status_counts.items():
            percentage = count / len(df) * 100
            bar = '█' * int(percentage / 2)
            md.append(f"- **{status}:** {count:,}건 ({percentage:.1f}%) {bar}")
        
        # 사이트별 인도 현황
        md.append("\n## 🏗️ 사이트별 인도 현황")
        for site in ['site_shu', 'site_mir', 'site_das', 'site_agi']:
            delivered = df[site].notna().sum()
            site_name = site.replace('site_', '').upper()
            md.append(f"- **{site_name}:** {delivered:,}건 인도 완료")
        
        # 지연 화물
        delayed = df[df['status'] == 'DELAYED']
        if len(delayed) > 0:
            md.append(f"\n## ⚠️ 지연 화물 ({len(delayed)}건)")
            for idx, row in delayed.head(10).iterrows():
                eta = row['eta']
                days_delayed = (pd.Timestamp.now() - eta).days if pd.notna(eta) else 0
                md.append(f"- **{row['case_no']}** - ETA: {eta.strftime('%Y-%m-%d') if pd.notna(eta) else 'N/A'} ({days_delayed}일 지연)")
        
        # 최근 인도 화물
        md.append("\n## 📦 최근 현장 인도 (최근 10건)")
        delivered_cargos = df[df['status'] == 'SITE_DELIVERED'].copy()
        if len(delivered_cargos) > 0:
            # 가장 최근 사이트 날짜 찾기
            site_dates = delivered_cargos[['site_shu', 'site_mir', 'site_das', 'site_agi']].apply(
                lambda row: row.dropna().max() if row.notna().any() else pd.NaT, axis=1
            )
            delivered_cargos['latest_delivery'] = site_dates
            delivered_cargos = delivered_cargos.sort_values('latest_delivery', ascending=False)
            
            for idx, row in delivered_cargos.head(10).iterrows():
                delivery_date = row['latest_delivery']
                md.append(f"- **{row['case_no']}** → {row['current_location']} ({delivery_date.strftime('%Y-%m-%d') if pd.notna(delivery_date) else 'N/A'})")
        
        # 이메일 연동 통계
        md.append("\n## 📧 이메일 연동 통계")
        with_emails = df[df['related_emails_count'] > 0]
        md.append(f"- **이메일 추적 가능:** {len(with_emails):,}건 ({len(with_emails)/len(df)*100:.1f}%)")
        md.append(f"- **이메일 없음:** {len(df) - len(with_emails):,}건")
        
        # 리드타임 분석
        md.append("\n## ⏱️ 리드타임 분석")
        shipping_days = []
        total_days = []
        for lead_time in df['lead_times']:
            if 'shipping_days' in lead_time:
                shipping_days.append(lead_time['shipping_days'])
            if 'total_days' in lead_time:
                total_days.append(lead_time['total_days'])
        
        if shipping_days:
            md.append(f"- **평균 해상운송 일수:** {np.mean(shipping_days):.1f}일")
            md.append(f"- **최소/최대:** {min(shipping_days)}일 / {max(shipping_days)}일")
        
        if total_days:
            md.append(f"- **평균 총 리드타임:** {np.mean(total_days):.1f}일")
            md.append(f"- **최소/최대:** {min(total_days)}일 / {max(total_days)}일")
        
        md.append("\n---")
        md.append("\n*Generated by HVDC Cargo Tracking System*")
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(md))
        
        print(f"   ✅ 리포트 저장 완료")
    
    def export_to_excel(self, output_file: str = "hvdc_cargo_tracking.xlsx"):
        """Excel 추적 파일 생성"""
        print(f"\n💾 Excel 추적 파일 생성: {output_file}")
        
        df = pd.DataFrame(self.tracking_data)
        
        # 날짜 컬럼 포맷팅
        date_cols = ['etd', 'atd', 'eta', 'ata', 'site_shu', 'site_mir', 'site_das', 'site_agi']
        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col]).dt.strftime('%Y-%m-%d')
        
        # 리드타임 컬럼 확장
        df['shipping_days'] = df['lead_times'].apply(lambda x: x.get('shipping_days', '') if isinstance(x, dict) else '')
        df['total_days'] = df['lead_times'].apply(lambda x: x.get('total_days', '') if isinstance(x, dict) else '')
        
        # 컬럼 제거
        df = df.drop(['lead_times', 'latest_email'], axis=1, errors='ignore')
        
        # Excel 저장
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Cargo Tracking', index=False)
            
            # 요약 시트
            summary_df = df.groupby('status_kr').agg({
                'case_no': 'count',
                'pkg': 'sum',
                'cbm': 'sum',
                'gwt': 'sum'
            }).reset_index()
            summary_df.columns = ['상태', '화물 수', '총 PKG', '총 CBM', '총 중량(KG)']
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        print(f"   ✅ Excel 파일 저장 완료")
    
    def generate_ontology_ttl(self, output_file: str = "hvdc_cargo_tracking_ontology.ttl"):
        """온톨로지 TTL 파일 생성"""
        print(f"\n📝 온톨로지 TTL 생성: {output_file}")
        
        ttl_lines = []
        
        # 네임스페이스
        ttl_lines.append("@prefix cargo: <http://samsung.com/hvdc/cargo#> .")
        ttl_lines.append("@prefix ex: <http://samsung.com/project-logistics#> .")
        ttl_lines.append("@prefix xsd: <http://www.w3.org/2001/XMLSchema#> .")
        ttl_lines.append("")
        
        # 클래스 정의
        ttl_lines.append("# Cargo Tracking Classes")
        ttl_lines.append("cargo:Cargo a rdfs:Class ;")
        ttl_lines.append('    rdfs:label "화물"@ko .')
        ttl_lines.append("")
        
        ttl_lines.append("cargo:CargoStatus a rdfs:Class ;")
        ttl_lines.append('    rdfs:label "화물 상태"@ko .')
        ttl_lines.append("")
        
        # 속성 정의
        ttl_lines.append("cargo:hasStatus a rdf:Property ;")
        ttl_lines.append('    rdfs:domain cargo:Cargo ;')
        ttl_lines.append('    rdfs:range cargo:CargoStatus .')
        ttl_lines.append("")
        
        ttl_lines.append("cargo:hasETA a rdf:Property ;")
        ttl_lines.append('    rdfs:domain cargo:Cargo ;')
        ttl_lines.append('    rdfs:range xsd:dateTime .')
        ttl_lines.append("")
        
        ttl_lines.append("cargo:hasATA a rdf:Property ;")
        ttl_lines.append('    rdfs:domain cargo:Cargo ;')
        ttl_lines.append('    rdfs:range xsd:dateTime .')
        ttl_lines.append("")
        
        ttl_lines.append("cargo:deliveredToSite a rdf:Property ;")
        ttl_lines.append('    rdfs:domain cargo:Cargo ;')
        ttl_lines.append('    rdfs:range ex:Site .')
        ttl_lines.append("")
        
        # 샘플 인스턴스 (50개)
        ttl_lines.append("# Cargo Instances (Sample 50)")
        for data in self.tracking_data[:50]:
            case_id = data['case_no'].replace('HVDC-ADOPT-', 'Cargo_')
            
            ttl_lines.append(f"cargo:{case_id} a cargo:Cargo ;")
            ttl_lines.append(f'    cargo:caseNumber "{data["case_no"]}" ;')
            ttl_lines.append(f'    cargo:hasStatus cargo:{data["status"]} ;')
            ttl_lines.append(f'    cargo:currentLocation "{data["current_location"]}" ;')
            
            if pd.notna(data.get('eta')):
                ttl_lines.append(f'    cargo:hasETA "{data["eta"]}"^^xsd:dateTime ;')
            if pd.notna(data.get('ata')):
                ttl_lines.append(f'    cargo:hasATA "{data["ata"]}"^^xsd:dateTime ;')
            
            # 사이트 인도 정보
            for site in ['shu', 'mir', 'das', 'agi']:
                site_date = data.get(f'site_{site}')
                if pd.notna(site_date):
                    ttl_lines.append(f'    cargo:deliveredToSite ex:Site_{site.upper()} ;')
                    ttl_lines.append(f'    cargo:deliveryDate_{site.upper()} "{site_date}"^^xsd:dateTime ;')
            
            ttl_lines[-1] = ttl_lines[-1].rstrip(' ;') + ' .'
            ttl_lines.append("")
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(ttl_lines))
        
        print(f"   ✅ TTL 파일 저장 완료")


def main():
    """메인 실행"""
    tracker = HVDCCargoTracker()
    
    # 데이터 로드
    if not tracker.load_data():
        return
    
    # 추적 데이터 처리
    tracker.process_tracking_data()
    
    # 리포트 생성
    tracker.generate_summary_report()
    
    # Excel 내보내기
    tracker.export_to_excel()
    
    # 온톨로지 생성
    tracker.generate_ontology_ttl()
    
    print("\n" + "="*70)
    print("✅ HVDC-ADOPT 화물 추적 시스템 실행 완료!")
    print("="*70)
    print("\n📂 생성된 파일:")
    print("   • hvdc_cargo_tracking_report.md - 추적 리포트")
    print("   • hvdc_cargo_tracking.xlsx - Excel 추적 파일")
    print("   • hvdc_cargo_tracking_ontology.ttl - 온톨로지 매핑")


if __name__ == "__main__":
    main()



