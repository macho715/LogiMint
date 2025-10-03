#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
리팩토링된 시스템으로 EMAIL 폴더 스캔 실행
"""

import sys
from pathlib import Path

# 현재 디렉토리를 Python 경로에 추가
sys.path.insert(0, str(Path(__file__).parent))

from hvdc.cli.scan_folders import scan_email_folders

def main():
    """메인 실행 함수"""
    try:
        print("🔍 EMAIL 폴더 스캔 시작...")
        
        # EMAIL 폴더 스캔
        results = scan_email_folders(Path('C:/Users/SAMSUNG/Documents/EMAIL'))
        
        print(f"✅ 스캔 완료!")
        print(f"📁 총 폴더: {results['total_folders']}개")
        print(f"📄 총 파일: {results['total_files']}개")
        print(f"🎯 고유 케이스: {results['unique_cases']}개")
        print(f"🏗️ 고유 사이트: {results['unique_sites']}개")
        print(f"📋 고유 LPO: {results['unique_lpos']}개")
        
        # 케이스 목록 출력 (처음 10개)
        if results['case_list']:
            print(f"\n📋 발견된 케이스 (처음 10개):")
            for i, case in enumerate(results['case_list'][:10]):
                print(f"  {i+1}. {case}")
        
        return 0
        
    except Exception as e:
        print(f"❌ 오류 발생: {e}")
        return 1

if __name__ == "__main__":
    exit(main())
