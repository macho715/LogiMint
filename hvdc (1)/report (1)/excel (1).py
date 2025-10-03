"""
엑셀 보고서 생성기 - 기존 5시트 구성 유지
"""
from __future__ import annotations
from pathlib import Path
from typing import List, Dict, Any
import pandas as pd
from datetime import datetime
from ..core.config import EXCEL_OUTDIR, EXCEL_FILENAME
from ..core.io import ensure_dir
from ..core.errors import IoError


def create_excel_report(data: List[Dict[str, Any]], output_path: Path = None) -> Path:
    """
    HVDC 이메일 데이터를 엑셀 보고서로 생성
    
    Args:
        data: 이메일 데이터 리스트
        output_path: 출력 파일 경로 (기본값: 타임스탬프 포함)
        
    Returns:
        Path: 생성된 엑셀 파일 경로
        
    Raises:
        IoError: 파일 생성 실패 시
    """
    try:
        # 출력 디렉토리 생성
        ensure_dir(EXCEL_OUTDIR)
        
        # 출력 파일 경로 설정
        if output_path is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = EXCEL_OUTDIR / f"hvdc_email_report_{timestamp}.xlsx"
        
        # 데이터프레임 생성
        df = pd.DataFrame(data)
        
        # 엑셀 파일 생성 (5시트 구성)
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 1. 전체 데이터
            df.to_excel(writer, sheet_name='전체_데이터', index=False)
            
            # 2. 사이트별 데이터
            if 'sites' in df.columns:
                site_data = _create_site_sheets(df)
                for site, site_df in site_data.items():
                    site_df.to_excel(writer, sheet_name=f'사이트_{site}', index=False)
            
            # 3. LPO별 데이터
            if 'lpos' in df.columns:
                lpo_data = _create_lpo_sheets(df)
                for lpo, lpo_df in lpo_data.items():
                    lpo_df.to_excel(writer, sheet_name=f'LPO_{lpo}', index=False)
            
            # 4. 긴급 데이터
            urgent_data = _create_urgent_sheets(df)
            if not urgent_data.empty:
                urgent_data.to_excel(writer, sheet_name='긴급_데이터', index=False)
            
            # 5. 배송 데이터
            delivery_data = _create_delivery_sheets(df)
            if not delivery_data.empty:
                delivery_data.to_excel(writer, sheet_name='배송_데이터', index=False)
        
        return output_path
        
    except Exception as e:
        raise IoError(f"엑셀 보고서 생성 실패: {output_path}") from e


def _create_site_sheets(df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """사이트별 데이터 시트 생성"""
    site_data = {}
    
    if 'sites' in df.columns:
        for site in df['sites'].dropna().unique():
            if site:
                site_df = df[df['sites'].str.contains(site, na=False)]
                site_data[site] = site_df
    
    return site_data


def _create_lpo_sheets(df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """LPO별 데이터 시트 생성"""
    lpo_data = {}
    
    if 'lpos' in df.columns:
        for lpo in df['lpos'].dropna().unique():
            if lpo:
                lpo_df = df[df['lpos'].str.contains(lpo, na=False)]
                lpo_data[lpo] = lpo_df
    
    return lpo_data


def _create_urgent_sheets(df: pd.DataFrame) -> pd.DataFrame:
    """긴급 데이터 시트 생성"""
    urgent_keywords = ['urgent', '긴급', 'emergency', 'asap']
    
    urgent_mask = df['subject'].str.contains('|'.join(urgent_keywords), case=False, na=False)
    return df[urgent_mask]


def _create_delivery_sheets(df: pd.DataFrame) -> pd.DataFrame:
    """배송 데이터 시트 생성"""
    delivery_keywords = ['delivery', '배송', 'shipping', 'transport']
    
    delivery_mask = df['subject'].str.contains('|'.join(delivery_keywords), case=False, na=False)
    return df[delivery_mask]
