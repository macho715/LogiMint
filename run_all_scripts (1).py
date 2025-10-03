#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HVDC 프로젝트 모든 스크립트 통합 실행기
"""

import subprocess
import sys
import time
import logging
from datetime import datetime
from pathlib import Path

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('hvdc_scripts_execution.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class HVDCScriptRunner:
    """HVDC 스크립트 통합 실행기"""
    
    def __init__(self):
        self.scripts = [
            {
                'name': 'email_folder_scanner',
                'file': 'email_folder_scanner.py',
                'description': 'EMAIL 폴더 스캔 및 이메일 데이터 추출',
                'priority': 1,
                'required': True
            },
            {
                'name': 'monitor_scan_progress',
                'file': 'monitor_scan_progress.py',
                'description': '스캔 진행상황 모니터링',
                'priority': 2,
                'required': False
            },
            {
                'name': 'email_ontology_mapper',
                'file': 'email_ontology_mapper.py',
                'description': '이메일 데이터 온톨로지 매핑',
                'priority': 3,
                'required': True
            },
            {
                'name': 'create_complete_email_excel',
                'file': 'create_complete_email_excel.py',
                'description': '종합 엑셀 보고서 생성',
                'priority': 4,
                'required': True
            },
            {
                'name': 'folder_title_mapper',
                'file': 'folder_title_mapper.py',
                'description': '폴더 제목 기반 매핑',
                'priority': 5,
                'required': False
            },
            {
                'name': 'simple_folder_analyzer',
                'file': 'simple_folder_analyzer.py',
                'description': '폴더 제목 간단 분석',
                'priority': 6,
                'required': False
            },
            {
                'name': 'hvdc_cargo_tracking_system',
                'file': 'hvdc_cargo_tracking_system.py',
                'description': '화물 추적 시스템',
                'priority': 7,
                'required': False
            },
            {
                'name': 'update_email_pattern_rules',
                'file': 'update_email_pattern_rules.py',
                'description': '이메일 패턴 규칙 업데이트',
                'priority': 8,
                'required': False
            },
            {
                'name': 'comprehensive_email_mapper',
                'file': 'comprehensive_email_mapper.py',
                'description': '종합 이메일 매핑 및 시각화',
                'priority': 9,
                'required': False
            },
            {
                'name': 'analysis_summary_report',
                'file': 'analysis_summary_report.py',
                'description': '분석 로직 종합 보고서 생성',
                'priority': 10,
                'required': False
            }
        ]
        
        self.results = []
        self.start_time = None
        self.end_time = None
    
    def check_dependencies(self):
        """의존성 확인"""
        logger.info("의존성 확인 중...")
        
        required_packages = [
            'pandas', 'numpy', 'matplotlib', 'seaborn', 
            'networkx', 'openpyxl', 'pathlib'
        ]
        
        missing_packages = []
        
        for package in required_packages:
            try:
                __import__(package)
                logger.info(f"✅ {package} 설치됨")
            except ImportError:
                missing_packages.append(package)
                logger.warning(f"❌ {package} 누락")
        
        if missing_packages:
            logger.error(f"누락된 패키지: {', '.join(missing_packages)}")
            logger.info("다음 명령어로 설치하세요:")
            logger.info(f"pip install {' '.join(missing_packages)}")
            return False
        
        logger.info("✅ 모든 의존성 확인 완료")
        return True
    
    def run_script(self, script_info):
        """개별 스크립트 실행"""
        script_name = script_info['name']
        script_file = script_info['file']
        description = script_info['description']
        
        logger.info(f"🚀 {script_name} 실행 시작: {description}")
        
        start_time = time.time()
        
        try:
            # Python 스크립트 실행
            result = subprocess.run(
                [sys.executable, script_file],
                capture_output=True,
                text=True,
                timeout=300  # 5분 타임아웃
            )
            
            end_time = time.time()
            execution_time = end_time - start_time
            
            if result.returncode == 0:
                logger.info(f"✅ {script_name} 실행 완료 ({execution_time:.2f}초)")
                self.results.append({
                    'script': script_name,
                    'status': 'SUCCESS',
                    'execution_time': execution_time,
                    'stdout': result.stdout,
                    'stderr': result.stderr
                })
                return True
            else:
                logger.error(f"❌ {script_name} 실행 실패 (코드: {result.returncode})")
                logger.error(f"오류: {result.stderr}")
                self.results.append({
                    'script': script_name,
                    'status': 'FAILED',
                    'execution_time': execution_time,
                    'stdout': result.stdout,
                    'stderr': result.stderr
                })
                return False
                
        except subprocess.TimeoutExpired:
            logger.error(f"⏰ {script_name} 실행 타임아웃 (5분 초과)")
            self.results.append({
                'script': script_name,
                'status': 'TIMEOUT',
                'execution_time': 300,
                'stdout': '',
                'stderr': 'Execution timeout after 5 minutes'
            })
            return False
            
        except Exception as e:
            logger.error(f"💥 {script_name} 실행 중 예외 발생: {str(e)}")
            self.results.append({
                'script': script_name,
                'status': 'ERROR',
                'execution_time': 0,
                'stdout': '',
                'stderr': str(e)
            })
            return False
    
    def run_all_scripts(self, required_only=False):
        """모든 스크립트 실행"""
        logger.info("🎯 HVDC 프로젝트 스크립트 통합 실행 시작")
        logger.info(f"실행 모드: {'필수 스크립트만' if required_only else '전체 스크립트'}")
        
        self.start_time = time.time()
        
        # 의존성 확인
        if not self.check_dependencies():
            logger.error("의존성 확인 실패. 실행을 중단합니다.")
            return False
        
        # 스크립트 실행
        success_count = 0
        total_count = 0
        
        for script_info in self.scripts:
            if required_only and not script_info['required']:
                logger.info(f"⏭️ {script_info['name']} 건너뜀 (선택적 스크립트)")
                continue
            
            total_count += 1
            
            if self.run_script(script_info):
                success_count += 1
            else:
                if script_info['required']:
                    logger.error(f"필수 스크립트 {script_info['name']} 실패. 실행을 중단합니다.")
                    break
        
        self.end_time = time.time()
        total_time = self.end_time - self.start_time
        
        # 결과 요약
        self.print_summary(success_count, total_count, total_time)
        
        return success_count == total_count
    
    def print_summary(self, success_count, total_count, total_time):
        """실행 결과 요약 출력"""
        logger.info("\n" + "="*60)
        logger.info("📊 HVDC 스크립트 실행 결과 요약")
        logger.info("="*60)
        
        logger.info(f"⏱️ 총 실행 시간: {total_time:.2f}초")
        logger.info(f"✅ 성공: {success_count}개")
        logger.info(f"❌ 실패: {total_count - success_count}개")
        logger.info(f"📈 성공률: {(success_count/total_count)*100:.1f}%")
        
        logger.info("\n📋 상세 결과:")
        for result in self.results:
            status_icon = {
                'SUCCESS': '✅',
                'FAILED': '❌',
                'TIMEOUT': '⏰',
                'ERROR': '💥'
            }.get(result['status'], '❓')
            
            logger.info(f"  {status_icon} {result['script']}: {result['status']} ({result['execution_time']:.2f}초)")
        
        # 실패한 스크립트 상세 정보
        failed_scripts = [r for r in self.results if r['status'] != 'SUCCESS']
        if failed_scripts:
            logger.info("\n🚨 실패한 스크립트 상세 정보:")
            for result in failed_scripts:
                logger.info(f"\n❌ {result['script']}:")
                if result['stderr']:
                    logger.info(f"  오류: {result['stderr']}")
    
    def run_interactive(self):
        """대화형 실행 모드"""
        print("\n🎯 HVDC 프로젝트 스크립트 통합 실행기")
        print("="*50)
        
        print("\n실행 옵션을 선택하세요:")
        print("1. 전체 스크립트 실행")
        print("2. 필수 스크립트만 실행")
        print("3. 개별 스크립트 실행")
        print("4. 종료")
        
        while True:
            try:
                choice = input("\n선택 (1-4): ").strip()
                
                if choice == '1':
                    self.run_all_scripts(required_only=False)
                    break
                elif choice == '2':
                    self.run_all_scripts(required_only=True)
                    break
                elif choice == '3':
                    self.run_individual_script()
                    break
                elif choice == '4':
                    print("프로그램을 종료합니다.")
                    break
                else:
                    print("잘못된 선택입니다. 1-4 중에서 선택하세요.")
                    
            except KeyboardInterrupt:
                print("\n프로그램을 종료합니다.")
                break
            except Exception as e:
                print(f"오류 발생: {str(e)}")
    
    def run_individual_script(self):
        """개별 스크립트 실행"""
        print("\n📋 실행 가능한 스크립트 목록:")
        for i, script in enumerate(self.scripts, 1):
            required_mark = " (필수)" if script['required'] else ""
            print(f"{i:2d}. {script['name']}{required_mark}")
            print(f"    {script['description']}")
        
        while True:
            try:
                choice = input(f"\n실행할 스크립트 번호 (1-{len(self.scripts)}, 0=종료): ").strip()
                
                if choice == '0':
                    break
                
                script_index = int(choice) - 1
                if 0 <= script_index < len(self.scripts):
                    script_info = self.scripts[script_index]
                    self.run_script(script_info)
                else:
                    print("잘못된 번호입니다.")
                    
            except ValueError:
                print("숫자를 입력하세요.")
            except KeyboardInterrupt:
                print("\n프로그램을 종료합니다.")
                break
            except Exception as e:
                print(f"오류 발생: {str(e)}")

def main():
    """메인 실행 함수"""
    runner = HVDCScriptRunner()
    
    # 명령행 인수 확인
    if len(sys.argv) > 1:
        if sys.argv[1] == '--all':
            runner.run_all_scripts(required_only=False)
        elif sys.argv[1] == '--required':
            runner.run_all_scripts(required_only=True)
        elif sys.argv[1] == '--interactive':
            runner.run_interactive()
        else:
            print("사용법:")
            print("  python run_all_scripts.py --all        # 전체 스크립트 실행")
            print("  python run_all_scripts.py --required   # 필수 스크립트만 실행")
            print("  python run_all_scripts.py --interactive # 대화형 실행")
    else:
        # 기본적으로 대화형 모드 실행
        runner.run_interactive()

if __name__ == "__main__":
    main()
