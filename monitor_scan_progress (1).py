#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EMAIL 폴더 스캔 진행상황 모니터링
"""

import os
import time
import json
from datetime import datetime
from pathlib import Path

def monitor_scan_progress():
    """스캔 진행상황 모니터링"""
    print("📧 EMAIL 폴더 스캔 진행상황 모니터링")
    print("=" * 50)
    
    # 로그 파일 확인
    log_file = "email_scan.log"
    if os.path.exists(log_file):
        print(f"📋 로그 파일: {log_file}")
        
        # 최근 로그 읽기
        with open(log_file, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        print(f"📊 총 로그 라인: {len(lines)}개")
        
        # 최근 10개 로그 출력
        print("\n🔍 최근 로그 (최근 10개):")
        for line in lines[-10:]:
            print(f"  {line.strip()}")
        
        # 폴더별 스캔 현황 분석
        folder_scan_count = 0
        outlook_skip_count = 0
        
        for line in lines:
            if "폴더 스캔 중:" in line:
                folder_scan_count += 1
            elif "Outlook 파일 스킵:" in line:
                outlook_skip_count += 1
        
        print(f"\n📈 스캔 통계:")
        print(f"  - 폴더 스캔: {folder_scan_count}개")
        print(f"  - Outlook 파일 스킵: {outlook_skip_count}개")
        
    else:
        print("❌ 로그 파일이 아직 생성되지 않았습니다.")
    
    # 결과 파일 확인
    result_files = [
        "email_scan_results_*.json",
        "email_folder_stats_*.csv", 
        "email_scan_report_*.md"
    ]
    
    print(f"\n📁 결과 파일 확인:")
    for pattern in result_files:
        import glob
        files = glob.glob(pattern)
        if files:
            for file in files:
                file_size = os.path.getsize(file)
                mod_time = datetime.fromtimestamp(os.path.getmtime(file))
                print(f"  ✅ {file} ({file_size:,} bytes, {mod_time.strftime('%H:%M:%S')})")
        else:
            print(f"  ⏳ {pattern} - 아직 생성되지 않음")

if __name__ == "__main__":
    monitor_scan_progress()
