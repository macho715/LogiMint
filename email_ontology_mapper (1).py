#!/usr/bin/env python3
"""
HVDC Email to Ontology Mapper
MACHO-GPT v3.4-mini - Email Intelligence System

기능:
- 이메일 데이터를 RDF/TTL 온톨로지로 변환
- 프로젝트 사이트, LPO, 벤더, 케이스 자동 추출
- HVDC 온톨로지 스키마 확장
"""

import pandas as pd
import re
from datetime import datetime
from pathlib import Path
from collections import defaultdict
from typing import Dict, List, Tuple, Set
import json


class HVDCEmailOntologyMapper:
    """이메일 데이터를 HVDC 온톨로지로 매핑"""
    
    def __init__(self, csv_path: str):
        self.csv_path = csv_path
        self.df = None
        self.ontology_data = []
        self.statistics = defaultdict(int)
        
        # 패턴 정의
        self.patterns = {
            'site': r'\b(DAS|AGI|MIR|MIRFA|GHALLAN)\b',
            'lpo': r'LPO[-\s]?(\d+)',
            'case': r'HVDC[-\s]([A-Z]+)[-\s]([A-Z]+)[-\s]([A-Z0-9\-]+)',
            'po': r'P[OJ]C?[-\s]LPO[-\s](\d+)',
            'invoice': r'(BAMF\d+|INV[-\s]?\d+)',
            'container': r'(CONTAINER|CONT|CNT)[-\s]?([A-Z0-9]+)',
            'delivery': r'(DELIVERY|DELIVER)',
            'urgent': r'(URGENT|IMMEDIATE|ASAP|CRITICAL)',
        }
        
        # 벤더 도메인 매핑
        self.vendor_domains = {}
        
        # 네임스페이스
        self.namespaces = {
            'ex': 'http://samsung.com/project-logistics#',
            'email': 'http://samsung.com/hvdc/email#',
            'vendor': 'http://samsung.com/hvdc/vendor#',
            'rdf': 'http://www.w3.org/1999/02/22-rdf-syntax-ns#',
            'rdfs': 'http://www.w3.org/2000/01/rdf-schema#',
            'xsd': 'http://www.w3.org/2001/XMLSchema#',
        }
    
    def load_data(self):
        """CSV 파일 로드"""
        print(f"📧 이메일 데이터 로드 중: {self.csv_path}")
        try:
            self.df = pd.read_csv(self.csv_path, encoding='utf-8-sig')
            print(f"   ✅ {len(self.df):,}개 이메일 로드 완료")
            print(f"   📅 기간: {self.df['received_time'].min()} ~ {self.df['received_time'].max()}")
            return True
        except Exception as e:
            print(f"   ❌ 오류: {e}")
            return False
    
    def extract_metadata(self, row: pd.Series) -> Dict:
        """이메일에서 메타데이터 추출"""
        subject = str(row['subject'])
        sender_email = str(row['sender_email'])
        
        metadata = {
            'email_id': f"email_{row.name}",
            'subject': subject,
            'sender_name': row['sender_name'],
            'sender_email': sender_email,
            'received_time': row['received_time'],
            'has_attachments': bool(row['has_attachments']),
            'folder': row['folder'],
        }
        
        # 사이트 추출
        site_match = re.search(self.patterns['site'], subject, re.IGNORECASE)
        if site_match:
            metadata['site'] = site_match.group(1).upper()
            self.statistics['emails_with_site'] += 1
        
        # LPO 번호 추출
        lpo_matches = re.findall(self.patterns['lpo'], subject, re.IGNORECASE)
        if lpo_matches:
            metadata['lpo_numbers'] = [f"LPO-{lpo}" for lpo in lpo_matches]
            self.statistics['emails_with_lpo'] += 1
        
        # 케이스 번호 추출
        case_match = re.search(self.patterns['case'], subject, re.IGNORECASE)
        if case_match:
            metadata['case_number'] = f"HVDC-{'-'.join(case_match.groups())}"
            self.statistics['emails_with_case'] += 1
        
        # 인보이스 추출
        invoice_match = re.search(self.patterns['invoice'], subject, re.IGNORECASE)
        if invoice_match:
            metadata['invoice_number'] = invoice_match.group(1)
            self.statistics['emails_with_invoice'] += 1
        
        # 컨테이너 추출
        container_match = re.search(self.patterns['container'], subject, re.IGNORECASE)
        if container_match:
            metadata['has_container'] = True
            self.statistics['emails_with_container'] += 1
        
        # 배송 관련
        if re.search(self.patterns['delivery'], subject, re.IGNORECASE):
            metadata['is_delivery'] = True
            self.statistics['delivery_emails'] += 1
        
        # 긴급 여부
        if re.search(self.patterns['urgent'], subject, re.IGNORECASE):
            metadata['is_urgent'] = True
            self.statistics['urgent_emails'] += 1
        
        # 벤더 도메인
        domain = sender_email.split('@')[-1] if '@' in sender_email else 'unknown'
        metadata['vendor_domain'] = domain
        self.vendor_domains[domain] = self.vendor_domains.get(domain, 0) + 1
        
        return metadata
    
    def classify_email_type(self, metadata: Dict) -> str:
        """이메일 타입 분류"""
        subject = metadata['subject'].lower()
        
        if 're:' in subject or 'fwd:' in subject:
            return 'Reply'
        elif 'lpo' in subject or 'purchase' in subject:
            return 'PurchaseOrder'
        elif 'delivery' in subject or 'deliver' in subject:
            return 'DeliveryNotification'
        elif 'invoice' in subject:
            return 'Invoice'
        elif 'quotation' in subject or 'rfq' in subject:
            return 'Quotation'
        elif 'approval' in subject or 'approve' in subject:
            return 'Approval'
        elif 'urgent' in subject or 'immediate' in subject:
            return 'Urgent'
        else:
            return 'General'
    
    def map_to_ontology(self):
        """온톨로지 매핑 실행"""
        print("\n🔄 온톨로지 매핑 시작...")
        
        for idx, row in self.df.iterrows():
            if idx % 1000 == 0:
                print(f"   진행률: {idx}/{len(self.df)} ({idx/len(self.df)*100:.1f}%)")
            
            # 메타데이터 추출
            metadata = self.extract_metadata(row)
            
            # 이메일 타입 분류
            email_type = self.classify_email_type(metadata)
            metadata['email_type'] = email_type
            
            # 온톨로지 데이터 생성
            self.ontology_data.append(metadata)
            self.statistics['total_processed'] += 1
        
        print(f"   ✅ {self.statistics['total_processed']:,}개 이메일 매핑 완료")
    
    def generate_ttl(self, output_file: str = "hvdc_emails_ontology.ttl"):
        """TTL 파일 생성"""
        print(f"\n📝 TTL 온톨로지 생성 중: {output_file}")
        
        ttl_lines = []
        
        # 네임스페이스 선언
        for prefix, uri in self.namespaces.items():
            ttl_lines.append(f"@prefix {prefix}: <{uri}> .")
        ttl_lines.append("")
        
        # 온톨로지 헤더
        ttl_lines.append("# =================================================================")
        ttl_lines.append("# HVDC Email Ontology - September 2025")
        ttl_lines.append("# MACHO-GPT v3.4-mini | Samsung C&T × ADNOC·DSV")
        ttl_lines.append(f"# Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        ttl_lines.append(f"# Total Emails: {len(self.ontology_data):,}")
        ttl_lines.append("# =================================================================")
        ttl_lines.append("")
        
        # 클래스 정의
        ttl_lines.append("# Email Classes")
        ttl_lines.append("email:Email a rdfs:Class ;")
        ttl_lines.append('    rdfs:label "이메일"@ko .')
        ttl_lines.append("")
        
        ttl_lines.append("email:PurchaseOrder a rdfs:Class ;")
        ttl_lines.append('    rdfs:subClassOf email:Email ;')
        ttl_lines.append('    rdfs:label "구매 주문"@ko .')
        ttl_lines.append("")
        
        ttl_lines.append("email:DeliveryNotification a rdfs:Class ;")
        ttl_lines.append('    rdfs:subClassOf email:Email ;')
        ttl_lines.append('    rdfs:label "배송 통지"@ko .')
        ttl_lines.append("")
        
        ttl_lines.append("vendor:Vendor a rdfs:Class ;")
        ttl_lines.append('    rdfs:label "벤더"@ko .')
        ttl_lines.append("")
        
        # 속성 정의
        ttl_lines.append("# Email Properties")
        ttl_lines.append("email:hasSubject a rdf:Property ;")
        ttl_lines.append('    rdfs:label "제목"@ko ;')
        ttl_lines.append('    rdfs:domain email:Email ;')
        ttl_lines.append('    rdfs:range xsd:string .')
        ttl_lines.append("")
        
        ttl_lines.append("email:hasSender a rdf:Property ;")
        ttl_lines.append('    rdfs:label "발신자"@ko ;')
        ttl_lines.append('    rdfs:domain email:Email ;')
        ttl_lines.append('    rdfs:range vendor:Vendor .')
        ttl_lines.append("")
        
        ttl_lines.append("email:receivedTime a rdf:Property ;")
        ttl_lines.append('    rdfs:label "수신 시간"@ko ;')
        ttl_lines.append('    rdfs:domain email:Email ;')
        ttl_lines.append('    rdfs:range xsd:dateTime .')
        ttl_lines.append("")
        
        ttl_lines.append("email:relatedToSite a rdf:Property ;")
        ttl_lines.append('    rdfs:label "관련 사이트"@ko ;')
        ttl_lines.append('    rdfs:domain email:Email ;')
        ttl_lines.append('    rdfs:range ex:Site .')
        ttl_lines.append("")
        
        ttl_lines.append("email:relatedToLPO a rdf:Property ;")
        ttl_lines.append('    rdfs:label "관련 LPO"@ko ;')
        ttl_lines.append('    rdfs:domain email:Email ;')
        ttl_lines.append('    rdfs:range xsd:string .')
        ttl_lines.append("")
        
        # 이메일 인스턴스 생성 (샘플 500개)
        ttl_lines.append("# Email Instances (Sample 500)")
        for i, data in enumerate(self.ontology_data[:500]):
            email_id = data['email_id']
            email_type = data['email_type']
            
            ttl_lines.append(f"email:{email_id} a email:{email_type} ;")
            ttl_lines.append(f'    email:hasSubject """{data["subject"]}""" ;')
            ttl_lines.append(f'    email:hasSender vendor:{self._sanitize_id(data["vendor_domain"])} ;')
            ttl_lines.append(f'    email:receivedTime "{data["received_time"]}"^^xsd:dateTime ;')
            ttl_lines.append(f'    email:hasAttachments "{str(data["has_attachments"]).lower()}"^^xsd:boolean ;')
            
            if 'site' in data:
                ttl_lines.append(f'    email:relatedToSite ex:Site_{data["site"]} ;')
            
            if 'lpo_numbers' in data:
                for lpo in data['lpo_numbers']:
                    ttl_lines.append(f'    email:relatedToLPO "{lpo}" ;')
            
            if 'case_number' in data:
                ttl_lines.append(f'    email:relatedToCase ex:Case_{self._sanitize_id(data["case_number"])} ;')
            
            if data.get('is_urgent'):
                ttl_lines.append(f'    email:isUrgent "true"^^xsd:boolean ;')
            
            # 마지막 줄 처리
            ttl_lines[-1] = ttl_lines[-1].rstrip(' ;') + ' .'
            ttl_lines.append("")
        
        # 벤더 인스턴스 (상위 20개)
        ttl_lines.append("# Top Vendors")
        top_vendors = sorted(self.vendor_domains.items(), key=lambda x: x[1], reverse=True)[:20]
        for domain, count in top_vendors:
            vendor_id = self._sanitize_id(domain)
            ttl_lines.append(f"vendor:{vendor_id} a vendor:Vendor ;")
            ttl_lines.append(f'    rdfs:label "{domain}" ;')
            ttl_lines.append(f'    vendor:emailCount {count} .')
            ttl_lines.append("")
        
        # 파일 저장
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(ttl_lines))
        
        print(f"   ✅ TTL 파일 저장 완료: {output_file}")
        print(f"   📦 총 {len(ttl_lines):,}줄 생성")
    
    def _sanitize_id(self, text: str) -> str:
        """ID로 사용 가능하도록 문자열 정리"""
        # 특수문자 제거 및 언더스코어로 치환
        text = re.sub(r'[^\w\-]', '_', text)
        text = re.sub(r'_+', '_', text)
        text = text.strip('_')
        return text
    
    def generate_analysis_report(self, output_file: str = "hvdc_email_analysis.md"):
        """분석 리포트 생성"""
        print(f"\n📊 분석 리포트 생성: {output_file}")
        
        md = []
        md.append("# HVDC 이메일 온톨로지 매핑 분석 리포트")
        md.append("\n**MACHO-GPT v3.4-mini** | Samsung C&T Logistics × ADNOC·DSV")
        md.append(f"\n**생성일:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        md.append(f"\n---\n")
        
        # 전체 통계
        md.append("## 📊 전체 통계")
        md.append(f"- **총 이메일 수:** {self.statistics['total_processed']:,}개")
        md.append(f"- **사이트 관련:** {self.statistics['emails_with_site']:,}개 ({self.statistics['emails_with_site']/self.statistics['total_processed']*100:.1f}%)")
        md.append(f"- **LPO 포함:** {self.statistics['emails_with_lpo']:,}개 ({self.statistics['emails_with_lpo']/self.statistics['total_processed']*100:.1f}%)")
        md.append(f"- **케이스 번호 포함:** {self.statistics['emails_with_case']:,}개 ({self.statistics['emails_with_case']/self.statistics['total_processed']*100:.1f}%)")
        md.append(f"- **인보이스 관련:** {self.statistics['emails_with_invoice']:,}개 ({self.statistics['emails_with_invoice']/self.statistics['total_processed']*100:.1f}%)")
        md.append(f"- **컨테이너 관련:** {self.statistics['emails_with_container']:,}개 ({self.statistics['emails_with_container']/self.statistics['total_processed']*100:.1f}%)")
        md.append(f"- **배송 관련:** {self.statistics['delivery_emails']:,}개 ({self.statistics['delivery_emails']/self.statistics['total_processed']*100:.1f}%)")
        md.append(f"- **긴급 이메일:** {self.statistics['urgent_emails']:,}개 ({self.statistics['urgent_emails']/self.statistics['total_processed']*100:.1f}%)")
        
        # 이메일 타입 분포
        md.append("\n## 📧 이메일 타입 분포")
        email_types = defaultdict(int)
        for data in self.ontology_data:
            email_types[data['email_type']] += 1
        
        for email_type, count in sorted(email_types.items(), key=lambda x: x[1], reverse=True):
            percentage = count / self.statistics['total_processed'] * 100
            md.append(f"- **{email_type}:** {count:,}개 ({percentage:.1f}%)")
        
        # 사이트별 분포
        md.append("\n## 🏗️ 사이트별 이메일 분포")
        sites = defaultdict(int)
        for data in self.ontology_data:
            if 'site' in data:
                sites[data['site']] += 1
        
        for site, count in sorted(sites.items(), key=lambda x: x[1], reverse=True):
            md.append(f"- **{site}:** {count:,}개")
        
        # 상위 벤더
        md.append("\n## 🏢 상위 벤더 (이메일 발송 기준)")
        top_vendors = sorted(self.vendor_domains.items(), key=lambda x: x[1], reverse=True)[:20]
        for i, (domain, count) in enumerate(top_vendors, 1):
            md.append(f"{i}. **{domain}:** {count:,}개")
        
        # 첨부파일 통계
        md.append("\n## 📎 첨부파일 통계")
        with_attachments = sum(1 for d in self.ontology_data if d['has_attachments'])
        md.append(f"- **첨부파일 있음:** {with_attachments:,}개 ({with_attachments/len(self.ontology_data)*100:.1f}%)")
        md.append(f"- **첨부파일 없음:** {len(self.ontology_data)-with_attachments:,}개")
        
        # 시간대별 분포
        md.append("\n## ⏰ 시간대별 이메일 분포")
        hours = defaultdict(int)
        for data in self.ontology_data:
            try:
                dt = datetime.fromisoformat(data['received_time'].replace('T', ' '))
                hours[dt.hour] += 1
            except:
                pass
        
        for hour in sorted(hours.keys()):
            count = hours[hour]
            bar = '█' * (count // 50)
            md.append(f"- **{hour:02d}시:** {count:,}개 {bar}")
        
        md.append("\n---")
        md.append("\n*Generated by HVDC Email Ontology Mapper*")
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(md))
        
        print(f"   ✅ 리포트 저장 완료: {output_file}")
    
    def export_to_json(self, output_file: str = "hvdc_emails_structured.json"):
        """구조화된 JSON 내보내기"""
        print(f"\n💾 JSON 데이터 내보내기: {output_file}")
        
        export_data = {
            'metadata': {
                'total_emails': len(self.ontology_data),
                'generated_at': datetime.now().isoformat(),
                'source_file': self.csv_path,
                'statistics': dict(self.statistics)
            },
            'emails': self.ontology_data[:1000],  # 샘플 1000개
            'vendors': [
                {'domain': domain, 'email_count': count}
                for domain, count in sorted(self.vendor_domains.items(), key=lambda x: x[1], reverse=True)[:50]
            ]
        }
        
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(export_data, f, ensure_ascii=False, indent=2)
        
        print(f"   ✅ JSON 저장 완료: {output_file}")


def main():
    """메인 실행 함수"""
    print("="*70)
    print("🚀 HVDC Email Ontology Mapper")
    print("   MACHO-GPT v3.4-mini | Samsung C&T Logistics")
    print("="*70)
    
    # 매퍼 초기화
    mapper = HVDCEmailOntologyMapper('emails_sep_full.csv')
    
    # 데이터 로드
    if not mapper.load_data():
        return
    
    # 온톨로지 매핑
    mapper.map_to_ontology()
    
    # TTL 생성
    mapper.generate_ttl('hvdc_emails_ontology.ttl')
    
    # 분석 리포트
    mapper.generate_analysis_report('hvdc_email_analysis.md')
    
    # JSON 내보내기
    mapper.export_to_json('hvdc_emails_structured.json')
    
    print("\n" + "="*70)
    print("✅ 모든 작업 완료!")
    print("="*70)
    print("\n📂 생성된 파일:")
    print("   • hvdc_emails_ontology.ttl - RDF/TTL 온톨로지 (500개 샘플)")
    print("   • hvdc_email_analysis.md - 분석 리포트")
    print("   • hvdc_emails_structured.json - 구조화된 JSON (1000개 샘플)")
    print("\n🔗 다음 단계:")
    print("   1. TTL 파일을 기존 HVDC 온톨로지와 통합")
    print("   2. SPARQL 쿼리로 이메일-프로젝트 관계 분석")
    print("   3. 시각화 도구로 이메일 네트워크 그래프 생성")


if __name__ == "__main__":
    main()



