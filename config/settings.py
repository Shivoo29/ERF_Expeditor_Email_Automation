"""Configuration settings for the ERF Email Automation system"""
import os
import json
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

class Config:
    """Main configuration class"""
    
    # Email settings
    SENDER_EMAIL = os.getenv('SENDER_EMAIL', 'your.email@lamresearch.com')
    SENDER_DISPLAY_NAME = os.getenv('SENDER_DISPLAY_NAME', 'Proto4Lab Team')
    LAM_DOMAIN = os.getenv('LAM_DOMAIN', 'lamresearch.com')
    TEST_MODE = os.getenv('TEST_MODE', 'True').lower() == 'true'
    
    # Outlook search settings
    SEARCH_OUTLOOK_CONTACTS = os.getenv('SEARCH_OUTLOOK_CONTACTS', 'True').lower() == 'true'
    
    # File paths
    DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'data')
    LOGS_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'logs')
    TEST_CONFIG_FILE = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'test_config.json')
    
    # Excel column mappings - UPDATED WITH YOUR ACTUAL COLUMNS
    REQUIRED_COLUMNS = [
        'Plnt', 'Ship-To-Plant', 'ERF Nr', 'Item', 'Entered by',
        'Material', 'Material Description', 'Unit', 'ERF Itm Qty',
        'ERF Sched Line Status', 'END', 'PO Due Date', 'Expeditor',
        'Expeditor Status',  'Expeditor Remarks'
    ]
    
    # Status filters
    TARGET_STATUSES = ['On order', 'Received']
    
    # Email settings
    EMAIL_SUBJECT_TEMPLATE = "ERF Status Update - {count} Items"
    
    @classmethod
    def ensure_directories(cls):
        """Ensure required directories exist"""
        os.makedirs(cls.DATA_DIR, exist_ok=True)
        os.makedirs(cls.LOGS_DIR, exist_ok=True)
    
    @classmethod
    def load_test_config(cls):
        """Load test configuration file"""
        if os.path.exists(cls.TEST_CONFIG_FILE):
            with open(cls.TEST_CONFIG_FILE, 'r') as f:
                return json.load(f)
        return None
    
    @classmethod
    def save_test_config(cls, config_data):
        """Save test configuration file"""
        with open(cls.TEST_CONFIG_FILE, 'w') as f:
            json.dump(config_data, f, indent=2)