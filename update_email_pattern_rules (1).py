#!/usr/bin/env python3
"""
이메일 제목에서 HVDC-ADOPT 케이스 번호 추출 규칙 개선
새로운 패턴 추가: PRL-XXX(HE-XXXX) 형태
"""

import pandas as pd
import re
from typing import List, Optional


def extract_case_numbers_enhanced(subject: str) -> List[str]:
    """
    개선된 케이스 번호 추출 (다중 패턴 지원)
    
    지원 패턴:
    1. HVDC-ADOPT-XXX-XXXX (기존)
    2. HVDC-XXX-XXX-XXXX (기존)
    3. (HE-XXXX) → HVDC-ADOPT-HE-XXXX (신규)
    4. (HE-XXXX-X) → HVDC-ADOPT-HE-XXXX-X (신규)
    5. (SIM-XXXX) → HVDC-ADOPT-SIM-XXXX (신규)
    6. (SCT-XXXX) → HVDC-ADOPT-SCT-XXXX (신규)
    """
    case_numbers = []
    
    # 패턴 1: 완전한 HVDC-ADOPT-XXX-XXXX 형태
    pattern1 = r'HVDC-ADOPT-([A-Z]+)-([A-Z0-9\-]+)'
    matches1 = re.findall(pattern1, subject, re.IGNORECASE)
    for match in matches1:
        case_numbers.append(f"HVDC-ADOPT-{match[0]}-{match[1]}".upper())
    
    # 패턴 2: HVDC-XXX-XXX-XXXX 형태 (ADOPT 없음)
    pattern2 = r'HVDC-([A-Z]+)-([A-Z]+)-([A-Z0-9\-]+)'
    matches2 = re.findall(pattern2, subject, re.IGNORECASE)
    for match in matches2:
        full_case = f"HVDC-{match[0]}-{match[1]}-{match[2]}".upper()
        if full_case not in case_numbers:  # 중복 방지
            case_numbers.append(full_case)
    
    # 패턴 3: 괄호 안의 약식 케이스 번호 (HE-XXXX), (SIM-XXXX), (SCT-XXXX) 등
    # 예: PRL-MIR-048-O1(HE-0427) 또는 (HE-0427, HE-0428)
    # 먼저 괄호 안의 전체 내용 추출
    pattern3_outer = r'\(([^\)]+)\)'
    outer_matches = re.findall(pattern3_outer, subject)
    
    for outer_match in outer_matches:
        # 괄호 안에서 개별 케이스 번호 추출
        # HE-0427, HE-0428 또는 HE-0427 형태
        pattern3_inner = r'([A-Z]+)-([0-9]+(?:-[0-9A-Z]+)?)'
        inner_matches = re.findall(pattern3_inner, outer_match, re.IGNORECASE)
        
        for match in inner_matches:
            vendor_code = match[0].upper()
            case_num = match[1]
            full_case = f"HVDC-ADOPT-{vendor_code}-{case_num}"
            if full_case not in case_numbers:
                case_numbers.append(full_case)
    
    # 패턴 4: 쉼표로 구분된 다중 케이스 번호
    # 예: (HE-0427, HE-0428) 또는 HE-0427,0428
    # 이미 패턴 3에서 개별적으로 추출됨
    
    # 패턴 5: PRL 코드에서 벤더 코드 추출
    # PRL-MIR-048, PRL-D-011 등에서 추가 정보 없음
    
    return case_numbers


def test_patterns():
    """테스트 케이스"""
    test_cases = [
        {
            'subject': 'RE: [Docu.Review] PRL-ZAK-031-A(HE-0504) Shims,Spare parts / AIR FREIGHT / GOT50300314',
            'expected': ['HVDC-ADOPT-HE-0504']
        },
        {
            'subject': 'RE: (URGENT) PRL-D-011-T-(HE-0499-3) // Delivery Request for 3150KVA CRT Transformer RTCC Panels - DAS station',
            'expected': ['HVDC-ADOPT-HE-0499-3']
        },
        {
            'subject': 'RE: [DELIVERY]PRL-MIR-048-O1, PRL-MIR-048-O2 (HE-0427, HE-0428) 400kV AC bus CT 20733/10 / CONTAINERS / GOT36800332-001,002',
            'expected': ['HVDC-ADOPT-HE-0427', 'HVDC-ADOPT-HE-0428']
        },
        {
            'subject': 'RE: [CUSTOMS] PRL-MIR-055-A(HE-0502) Insulators / AIR FREIGHT / GOT50300312',
            'expected': ['HVDC-ADOPT-HE-0502']
        },
        {
            'subject': 'RE: [HVDC-DAS] Best Way Equipment_SCT-19LT-PJC-LPO-1487_5 Units',
            'expected': []  # LPO는 케이스 번호가 아님
        },
        {
            'subject': '[HVDC-ADOPT-SIM-0088] Material Delivery',
            'expected': ['HVDC-ADOPT-SIM-0088']
        },
        {
            'subject': 'HVDC-AGI-SCT-0134 Project Update',
            'expected': ['HVDC-AGI-SCT-0134']
        }
    ]
    
    print("="*70)
    print("🧪 케이스 번호 추출 패턴 테스트")
    print("="*70)
    
    all_passed = True
    for i, test in enumerate(test_cases, 1):
        subject = test['subject']
        expected = test['expected']
        result = extract_case_numbers_enhanced(subject)
        
        passed = set(result) == set(expected)
        status = "✅ PASS" if passed else "❌ FAIL"
        
        print(f"\n테스트 {i}: {status}")
        print(f"제목: {subject[:80]}...")
        print(f"예상: {expected}")
        print(f"결과: {result}")
        
        if not passed:
            all_passed = False
    
    print("\n" + "="*70)
    if all_passed:
        print("✅ 모든 테스트 통과!")
    else:
        print("❌ 일부 테스트 실패")
    print("="*70)
    
    return all_passed


def update_excel_with_enhanced_patterns():
    """기존 Excel 파일을 개선된 패턴으로 업데이트"""
    print("\n📊 Excel 파일 업데이트 시작...")
    
    # 원본 파일 로드
    df = pd.read_excel('emails_sep_full_mapped.xlsx', sheet_name='전체이메일')
    print(f"   로드: {len(df):,}건")
    
    # 기존 케이스 번호 개수
    old_case_count = df['case_number'].notna().sum()
    print(f"   기존 케이스 번호: {old_case_count:,}건")
    
    # 개선된 패턴으로 재추출
    def extract_and_join(subject):
        cases = extract_case_numbers_enhanced(str(subject))
        return ', '.join(cases) if cases else None
    
    df['case_number_new'] = df['subject'].apply(extract_and_join)
    
    # 새로운 케이스 번호 개수
    new_case_count = df['case_number_new'].notna().sum()
    print(f"   새로운 케이스 번호: {new_case_count:,}건")
    print(f"   증가: +{new_case_count - old_case_count:,}건")
    
    # 기존 케이스 번호와 병합 (새로운 것 우선)
    df['case_number_final'] = df['case_number_new'].fillna(df['case_number'])
    final_count = df['case_number_final'].notna().sum()
    
    # 업데이트된 컬럼으로 교체
    df['case_number'] = df['case_number_final']
    df = df.drop(['case_number_new', 'case_number_final'], axis=1)
    
    # 새로 추출된 케이스 샘플 출력
    print("\n📋 새로 추출된 케이스 번호 샘플:")
    new_cases = df[df['case_number'].notna() & (df.index >= old_case_count)].head(10)
    if len(new_cases) > 0:
        for idx, row in new_cases.iterrows():
            print(f"   • {row['case_number'][:50]} - {row['subject'][:60]}...")
    
    # Excel 저장
    output_file = 'emails_sep_full_mapped_v2.xlsx'
    print(f"\n💾 업데이트된 파일 저장: {output_file}")
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # 전체 데이터
        df.to_excel(writer, sheet_name='전체이메일', index=False)
        
        # 케이스별 (업데이트된 데이터)
        df_case = df[df['case_number'].notna()].copy()
        
        # 케이스 번호 분리 (쉼표로 구분된 경우)
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
            'sender_domain': lambda x: ', '.join(x.unique()[:3])
        }).reset_index()
        case_summary.columns = ['케이스번호', '이메일수', '최초수신', '최근수신', '관련벤더']
        case_summary = case_summary.sort_values('이메일수', ascending=False)
        case_summary.to_excel(writer, sheet_name='케이스별_개선', index=False)
        
        print(f"   ✅ 시트 1: 전체이메일 ({len(df):,}건)")
        print(f"   ✅ 시트 2: 케이스별_개선 ({len(case_summary):,}건)")
        
        # 요약 통계
        summary_data = {
            '항목': [
                '총 이메일 수',
                '케이스 번호 있음 (기존)',
                '케이스 번호 있음 (개선)',
                '증가분',
                '증가율'
            ],
            '값': [
                f"{len(df):,}건",
                f"{old_case_count:,}건",
                f"{final_count:,}건",
                f"+{final_count - old_case_count:,}건",
                f"+{(final_count - old_case_count)/old_case_count*100:.1f}%"
            ]
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='개선통계', index=False)
        print(f"   ✅ 시트 3: 개선통계")
    
    print(f"\n✅ 업데이트 완료!")
    print(f"   파일: {output_file}")
    print(f"   케이스 번호 추출: {old_case_count:,}건 → {final_count:,}건 (+{final_count-old_case_count:,}건)")
    
    return output_file


def main():
    """메인 실행"""
    # 1. 패턴 테스트
    test_passed = test_patterns()
    
    if not test_passed:
        print("\n⚠️  테스트 실패 - Excel 업데이트를 건너뜁니다.")
        return
    
    # 2. Excel 파일 업데이트
    print("\n" + "="*70)
    output_file = update_excel_with_enhanced_patterns()
    
    print("\n" + "="*70)
    print("📂 생성된 파일:")
    print(f"   {output_file}")
    print("="*70)


if __name__ == "__main__":
    main()

