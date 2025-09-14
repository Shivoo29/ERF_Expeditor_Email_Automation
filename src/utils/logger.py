# src/utils/logger.py
"""Logging utility without Unicode emojis to avoid Windows console issues"""
import logging
import os
import sys
from datetime import datetime
from config.settings import Config

def setup_logger(name: str, level: int = logging.INFO) -> logging.Logger:
    """Set up logger with file and console handlers (no Unicode emojis)"""
    
    # Ensure logs directory exists
    Config.ensure_directories()
    
    # Create logger
    logger = logging.getLogger(name)
    logger.setLevel(level)
    
    # Avoid duplicate handlers
    if logger.handlers:
        return logger
    
    # Create formatter without Unicode characters
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    # File handler
    log_file = os.path.join(
        Config.LOGS_DIR, 
        f'erf_automation_{datetime.now().strftime("%Y%m%d")}.log'
    )
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(level)
    file_handler.setFormatter(formatter)
    
    # Console handler with UTF-8 encoding for Windows
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(level)
    console_handler.setFormatter(formatter)
    
    # Add handlers
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger