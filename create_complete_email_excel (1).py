#!/usr/bin/env python3
"""
완전한 이메일 Excel 파일 생성 (모든 시트 포함)
개선된 케이스 번호 추출 패턴 적용
"""

import pandas as pd
import re
from datetime import datetime
from pathlib import Path


def extract_case_numbers_enhanced(subject: str):
    """개선된 케이스 번호 추출 (JPTW/GRM 패턴 포함)"""
    case_numbers = []
    subject_str = str(subject)
    
    # 패턴 1: 완전한 HVDC-ADOPT-XXX-XXXX
    pattern1 = r'HVDC-ADOPT-([A-Z]+)-([A-Z0-9\-]+)'
    matches1 = re.findall(pattern1, subject_str, re.IGNORECASE)
    for match in matches1:
        case_numbers.append(f"HVDC-ADOPT-{match[0]}-{match[1]}".upper())
    
    # 패턴 2: HVDC-XXX-XXX-XXXX
    pattern2 = r'HVDC-([A-Z]+)-([A-Z]+)-([A-Z0-9\-]+)'
    matches2 = re.findall(pattern2, subject_str, re.IGNORECASE)
    for match in matches2:
        full_case = f"HVDC-{match[0]}-{match[1]}-{match[2]}".upper()
        if full_case not in case_numbers:
            case_numbers.append(full_case)
    
    # 패턴 3: 괄호 안의 약식 (HE-XXXX), (SIM-XXXX) 등
    pattern3_outer = r'\(([^\)]+)\)'
    outer_matches = re.findall(pattern3_outer, subject_str)
    
    for outer_match in outer_matches:
        pattern3_inner = r'([A-Z]+)-([0-9]+(?:-[0-9A-Z]+)?)'
        inner_matches = re.findall(pattern3_inner, outer_match, re.IGNORECASE)
        
        for match in inner_matches:
            vendor_code = match[0].upper()
            case_num = match[1]
            full_case = f"HVDC-ADOPT-{vendor_code}-{case_num}"
            if full_case not in case_numbers:
                case_numbers.append(full_case)
    
    # 패턴 4: JPTW-XX / GRM-XXX 형태 (AGI 사이트)
    pattern4 = r'\[HVDC-AGI\].*?(JPTW-(\d+))\s*/\s*(GRM-(\d+))'
    matches4 = re.findall(pattern4, subject_str, re.IGNORECASE)
    for match in matches4:
        jptw_num = match[1]
        grm_num = match[3]
        full_case = f"HVDC-AGI-JPTW{jptw_num}-GRM{grm_num}".upper()
        if full_case not in case_numbers:
            case_numbers.append(full_case)
    
    # 패턴 5: 콜론 뒤 완성된 케이스 번호 (HVDC-AGI-JPTW71-GRM65)
    pattern5 = r':\s*([A-Z]+-[A-Z]+-[A-Z]+\d+-[A-Z]+\d+)'
    matches5 = re.findall(pattern5, subject_str, re.IGNORECASE)
    for match in matches5:
        clean_case = re.sub(r'\(.*?\)', '', match).strip().upper()
        if clean_case not in case_numbers:
            case_numbers.append(clean_case)
    
    return ', '.join(case_numbers) if case_numbers else None


def create_complete_excel():
    """완전한 Excel 파일 생성"""
    print("="*70)
    print("📧 완전한 이메일 Excel 파일 생성")
    print("   개선된 케이스 번호 추출 + 전체 시트")
    print("="*70)
    
    # CSV 로드
    print("\n📂 CSV 파일 로드: emails_sep_full.csv")
    df = pd.read_csv('emails_sep_full.csv', encoding='utf-8-sig')
    print(f"   ✅ {len(df):,}건 로드 완료")
    
    # 날짜 파싱
    df['received_time'] = pd.to_datetime(df['received_time'])
    df['date'] = df['received_time'].dt.date
    df['time'] = df['received_time'].dt.time
    df['hour'] = df['received_time'].dt.hour
    df['weekday'] = df['received_time'].dt.day_name()
    
    # 도메인 추출
    df['sender_domain'] = df['sender_email'].apply(
        lambda x: x.split('@')[-1] if '@' in str(x) else 'unknown'
    )
    
    # 개선된 패턴으로 추출
    print("\n🔄 데이터 추출 중...")
    
    # 사이트 추출
    def extract_site(subject):
        match = re.search(r'\b(DAS|AGI|MIR|MIRFA|GHALLAN)\b', str(subject), re.IGNORECASE)
        return match.group(1).upper() if match else None
    
    # LPO 추출
    def extract_lpo(subject):
        matches = re.findall(r'LPO[-\s]?(\d+)', str(subject), re.IGNORECASE)
        return ', '.join([f"LPO-{m}" for m in matches]) if matches else None
    
    # 케이스 번호 추출 (개선된 버전)
    df['site'] = df['subject'].apply(extract_site)
    df['lpo_numbers'] = df['subject'].apply(extract_lpo)
    df['case_number'] = df['subject'].apply(extract_case_numbers_enhanced)
    
    # 이메일 타입 분류
    def classify_email_type(subject):
        subject_lower = str(subject).lower()
        if 're:' in subject_lower or 'fwd:' in subject_lower:
            return 'Reply'
        elif 'lpo' in subject_lower or 'purchase' in subject_lower:
            return 'PurchaseOrder'
        elif 'delivery' in subject_lower or 'deliver' in subject_lower:
            return 'DeliveryNotification'
        elif 'invoice' in subject_lower:
            return 'Invoice'
        elif 'quotation' in subject_lower or 'rfq' in subject_lower:
            return 'Quotation'
        elif 'approval' in subject_lower or 'approve' in subject_lower:
            return 'Approval'
        elif 'urgent' in subject_lower or 'immediate' in subject_lower:
            return 'Urgent'
        else:
            return 'General'
    
    df['email_type'] = df['subject'].apply(classify_email_type)
    df['is_urgent'] = df['subject'].str.contains('URGENT|IMMEDIATE|CRITICAL', case=False, na=False)
    df['is_delivery'] = df['subject'].str.contains('DELIVERY|DELIVER', case=False, na=False)
    df['has_container'] = df['subject'].str.contains('CONTAINER|CONT|CNT', case=False, na=False)
    
    print(f"   ✅ 사이트 관련: {df['site'].notna().sum():,}건")
    print(f"   ✅ LPO 관련: {df['lpo_numbers'].notna().sum():,}건")
    print(f"   ✅ 케이스 번호: {df['case_number'].notna().sum():,}건")
    
    # Excel 저장
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = f'emails_sep_full_complete_v{timestamp}.xlsx'
    print(f"\n💾 Excel 파일 생성: {output_file}")
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        
        # 시트 1: 전체이메일
        export_cols = [
            'folder', 'sender_name', 'sender_email', 'sender_domain',
            'received_time', 'date', 'time', 'hour', 'weekday',
            'subject', 'has_attachments',
            'site', 'lpo_numbers', 'case_number',
            'email_type', 'is_urgent', 'is_delivery', 'has_container'
        ]
        df_export = df[export_cols].copy()
        df_export.to_excel(writer, sheet_name='전체이메일', index=False)
        print(f"   ✅ 시트 1: 전체이메일 ({len(df_export):,}건)")
        
        # 시트 2: 사이트별
        df_site = df[df['site'].notna()].copy()
        df_site = df_site.sort_values(['site', 'received_time'], ascending=[True, False])
        df_site[export_cols].to_excel(writer, sheet_name='사이트별', index=False)
        print(f"   ✅ 시트 2: 사이트별 ({len(df_site):,}건)")
        
        # 시트 3: LPO관련
        df_lpo = df[df['lpo_numbers'].notna()].copy()
        df_lpo = df_lpo.sort_values('received_time', ascending=False)
        df_lpo[export_cols].to_excel(writer, sheet_name='LPO관련', index=False)
        print(f"   ✅ 시트 3: LPO관련 ({len(df_lpo):,}건)")
        
        # 시트 4: 긴급이메일
        df_urgent = df[df['is_urgent'] == True].copy()
        df_urgent = df_urgent.sort_values('received_time', ascending=False)
        df_urgent[['sender_name', 'sender_email', 'received_time', 'subject', 
                   'site', 'case_number', 'lpo_numbers']].to_excel(
            writer, sheet_name='긴급이메일', index=False
        )
        print(f"   ✅ 시트 4: 긴급이메일 ({len(df_urgent):,}건)")
        
        # 시트 5: 배송관련
        df_delivery = df[df['is_delivery'] == True].copy()
        df_delivery = df_delivery.sort_values('received_time', ascending=False)
        df_delivery[['sender_name', 'sender_email', 'received_time', 'subject', 
                     'site', 'case_number', 'lpo_numbers']].to_excel(
            writer, sheet_name='배송관련', index=False
        )
        print(f"   ✅ 시트 5: 배송관련 ({len(df_delivery):,}건)")
        
        # 시트 6: 벤더별통계
        vendor_stats = df.groupby('sender_domain').agg({
            'sender_email': 'count',
            'has_attachments': 'sum',
            'site': lambda x: x.notna().sum(),
            'lpo_numbers': lambda x: x.notna().sum(),
            'case_number': lambda x: x.notna().sum()
        }).reset_index()
        vendor_stats.columns = ['벤더도메인', '총이메일수', '첨부파일있음', 
                               '사이트관련', 'LPO관련', '케이스관련']
        vendor_stats = vendor_stats.sort_values('총이메일수', ascending=False)
        vendor_stats.to_excel(writer, sheet_name='벤더별통계', index=False)
        print(f"   ✅ 시트 6: 벤더별통계 ({len(vendor_stats):,}개 벤더)")
        
        # 시트 7: 일자별통계
        daily_stats = df.groupby('date').agg({
            'sender_email': 'count',
            'has_attachments': 'sum',
            'is_urgent': 'sum',
            'is_delivery': 'sum',
            'site': lambda x: x.notna().sum(),
            'case_number': lambda x: x.notna().sum()
        }).reset_index()
        daily_stats.columns = ['날짜', '총이메일', '첨부파일', '긴급', 
                              '배송', '사이트관련', '케이스관련']
        daily_stats = daily_stats.sort_values('날짜', ascending=False)
        daily_stats.to_excel(writer, sheet_name='일자별통계', index=False)
        print(f"   ✅ 시트 7: 일자별통계 ({len(daily_stats):,}일)")
        
        # 시트 8: 시간대별통계
        hourly_stats = df.groupby('hour').agg({
            'sender_email': 'count',
            'is_urgent': 'sum',
            'is_delivery': 'sum'
        }).reset_index()
        hourly_stats.columns = ['시간대', '총이메일', '긴급', '배송']
        hourly_stats.to_excel(writer, sheet_name='시간대별통계', index=False)
        print(f"   ✅ 시트 8: 시간대별통계 (24시간)")
        
        # 시트 9: 케이스별
        df_case = df[df['case_number'].notna()].copy()
        
        # 쉼표로 구분된 케이스 분리
        case_rows = []
        for idx, row in df_case.iterrows():
            cases = str(row['case_number']).split(', ')
            for case in cases:
                case_row = row.copy()
                case_row['case_number'] = case.strip()
                case_rows.append(case_row)
        
        df_case_expanded = pd.DataFrame(case_rows)
        
        case_summary = df_case_expanded.groupby('case_number').agg({
            'sender_email': 'count',
            'received_time': ['min', 'max'],
            'sender_domain': lambda x: ', '.join(x.unique()[:3]),
            'is_urgent': 'sum',
            'is_delivery': 'sum'
        }).reset_index()
        case_summary.columns = ['케이스번호', '이메일수', '최초수신', '최근수신', 
                               '관련벤더', '긴급건', '배송건']
        case_summary = case_summary.sort_values('이메일수', ascending=False)
        case_summary.to_excel(writer, sheet_name='케이스별', index=False)
        print(f"   ✅ 시트 9: 케이스별 ({len(case_summary):,}건)")
        
        # 시트 10: 요약통계
        summary_data = {
            '항목': [
                '총 이메일 수',
                '기간',
                '발신자 수',
                '벤더 수',
                '첨부파일 있음',
                '사이트 관련',
                'LPO 관련',
                '케이스 번호 있음',
                '긴급 이메일',
                '배송 관련',
                '컨테이너 관련',
                '---구분선---',
                '이메일 타입 - Reply',
                '이메일 타입 - General',
                '이메일 타입 - PurchaseOrder',
                '이메일 타입 - DeliveryNotification',
                '이메일 타입 - Urgent',
                '---사이트별---',
                'DAS',
                'AGI',
                'MIR',
                'MIRFA',
                'GHALLAN'
            ],
            '값': [
                f"{len(df):,}건",
                f"{df['date'].min()} ~ {df['date'].max()}",
                f"{df['sender_email'].nunique():,}명",
                f"{df['sender_domain'].nunique():,}개",
                f"{df['has_attachments'].sum():,}건 ({df['has_attachments'].sum()/len(df)*100:.1f}%)",
                f"{df['site'].notna().sum():,}건 ({df['site'].notna().sum()/len(df)*100:.1f}%)",
                f"{df['lpo_numbers'].notna().sum():,}건 ({df['lpo_numbers'].notna().sum()/len(df)*100:.1f}%)",
                f"{df['case_number'].notna().sum():,}건 ({df['case_number'].notna().sum()/len(df)*100:.1f}%)",
                f"{df['is_urgent'].sum():,}건 ({df['is_urgent'].sum()/len(df)*100:.1f}%)",
                f"{df['is_delivery'].sum():,}건 ({df['is_delivery'].sum()/len(df)*100:.1f}%)",
                f"{df['has_container'].sum():,}건 ({df['has_container'].sum()/len(df)*100:.1f}%)",
                '---',
                f"{(df['email_type']=='Reply').sum():,}건 ({(df['email_type']=='Reply').sum()/len(df)*100:.1f}%)",
                f"{(df['email_type']=='General').sum():,}건 ({(df['email_type']=='General').sum()/len(df)*100:.1f}%)",
                f"{(df['email_type']=='PurchaseOrder').sum():,}건 ({(df['email_type']=='PurchaseOrder').sum()/len(df)*100:.1f}%)",
                f"{(df['email_type']=='DeliveryNotification').sum():,}건 ({(df['email_type']=='DeliveryNotification').sum()/len(df)*100:.1f}%)",
                f"{(df['email_type']=='Urgent').sum():,}건 ({(df['email_type']=='Urgent').sum()/len(df)*100:.1f}%)",
                '---',
                f"{(df['site']=='DAS').sum():,}건",
                f"{(df['site']=='AGI').sum():,}건",
                f"{(df['site']=='MIR').sum():,}건",
                f"{(df['site']=='MIRFA').sum():,}건",
                f"{(df['site']=='GHALLAN').sum():,}건"
            ]
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='요약통계', index=False)
        print(f"   ✅ 시트 10: 요약통계")
    
    # 파일 정보
    file_size = Path(output_file).stat().st_size / 1024 / 1024
    print(f"\n✅ Excel 파일 생성 완료!")
    print(f"   파일명: {output_file}")
    print(f"   총 시트: 10개")
    print(f"   파일 크기: {file_size:.2f} MB")
    
    # 통계 출력
    print(f"\n📊 최종 통계:")
    print(f"   총 이메일: {len(df):,}건")
    print(f"   사이트별: {df['site'].notna().sum():,}건")
    print(f"   LPO관련: {df['lpo_numbers'].notna().sum():,}건")
    print(f"   케이스번호: {df['case_number'].notna().sum():,}건")
    print(f"   긴급: {df['is_urgent'].sum():,}건")
    print(f"   배송: {df['is_delivery'].sum():,}건")
    
    return output_file


if __name__ == "__main__":
    output_file = create_complete_excel()
    
    print("\n" + "="*70)
    print("📂 Excel 파일 열기:")
    print(f"   {Path(output_file).absolute()}")
    print("="*70)

