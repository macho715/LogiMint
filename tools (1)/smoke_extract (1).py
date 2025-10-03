#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
임시 도우미 스크립트 - 새 추출기 스모크 테스트
"""

from pathlib import Path
import sys

# 프로젝트 루트를 Python 경로에 추가
sys.path.insert(0, str(Path(__file__).parent.parent))

from hvdc.extractors.case import extract_cases
from hvdc.extractors.site import extract_sites
from hvdc.extractors.lpo import extract_lpos

def main():
    """새 추출기 스모크 테스트"""
    print("🔍 HVDC 추출기 스모크 테스트")
    print("=" * 50)
    
    samples = [
        "RE: HVDC-ADOPT-HE-0476 - JPTW-71 / GRM-123",
        "PRL-O-046-O4(HE-0486) - PO 5001005009 AGI",
        "HVDC-DAS-HMU-MOSB-0164 - DAS SITE",
        "RE: [HVDC-AGI] JPTW-71 / GRM-065 – Aggregate 20mm (760 Ton) – 11th Sep",
        "HVDC-AGI-JPTW71-GRM65(760TN)",
    ]
    
    for i, sample in enumerate(samples, 1):
        print(f"\n📝 샘플 {i}: {sample}")
        
        try:
            cases = extract_cases(sample)
            sites = extract_sites(sample)
            lpos = extract_lpos(sample)
            
            print(f"  🎯 케이스: {[h['value'] for h in cases]}")
            print(f"  🏗️ 사이트: {sites}")
            print(f"  📋 LPO: {lpos}")
            
        except Exception as e:
            print(f"  ❌ 오류: {e}")
    
    print("\n✅ 스모크 테스트 완료")

if __name__ == "__main__":
    main()
