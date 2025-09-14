"""Validation utilities"""
import os
import pandas as pd
from typing import List, Tuple
from config.settings import Config

def validate_excel_file(file_path: str) -> Tuple[bool, str]:
    """Validate if Excel file exists and is accessible"""
    if not os.path.exists(file_path):
        return False, f"File not found: {file_path}"
    
    if not file_path.lower().endswith(('.xlsx', '.xls')):
        return False, "File must be an Excel file (.xlsx or .xls)"
    
    try:
        # Try to read first few rows
        pd.read_excel(file_path, nrows=1)
        return True, "File is valid"
    except Exception as e:
        return False, f"Cannot read Excel file: {str(e)}"

def validate_dataframe_columns(df: pd.DataFrame) -> Tuple[bool, str, List[str]]:
    """Validate if DataFrame has required columns"""
    df_columns = list(df.columns)
    missing_columns = [col for col in Config.REQUIRED_COLUMNS if col not in df_columns]
    
    if missing_columns:
        return False, f"Missing columns: {missing_columns}", missing_columns
    
    return True, "All required columns present", []
