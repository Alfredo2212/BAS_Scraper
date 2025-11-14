"""
Settings and configuration for OJK Scraper
"""

import os
from pathlib import Path

# Base paths
PROJECT_ROOT = Path(__file__).parent.parent
OUTPUT_DIR = PROJECT_ROOT / "output"
LOGS_DIR = PROJECT_ROOT / "logs"

# Create directories if they don't exist
OUTPUT_DIR.mkdir(exist_ok=True)
LOGS_DIR.mkdir(exist_ok=True)


class OJKConfig:
    """Configuration for OJK website scraping"""
    
    # Base URL
    BASE_URL = "https://ojk.go.id/id/kanal/perbankan/data-dan-statistik/laporan-keuangan-perbankan/Default.aspx"
    
    # Timeouts (in seconds)
    PAGE_LOAD_TIMEOUT = 30
    ELEMENT_WAIT_TIMEOUT = 30
    POSTBACK_WAIT_TIMEOUT = 30
    
    # Delays (in seconds) - to avoid rate limiting
    DELAY_BETWEEN_REQUESTS = 2.0
    DELAY_AFTER_POSTBACK = 3.0
    RANDOM_DELAY_RANGE = (0, 1)  # Additional random delay
    
    # Selenium settings
    HEADLESS_MODE = False  # Set to True for headless mode
    WINDOW_SIZE = (1920, 1080)
    
    # User agents for rotation
    USER_AGENTS = [
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    ]
    
    # XPath selectors (using contains() for dynamic IDs)
    SELECTORS = {
        # Tab selector: Target the clickable button element (x-tab-button) that contains the text
        # Alternative simpler selector: //span[contains(@class, 'x-tab-inner') and contains(text(), 'BPR Konvensional')]
        'tab_bpr_konvensional': "//span[contains(@class, 'x-tab-inner') and contains(text(), 'BPR Konvensional')]/ancestor::span[contains(@class, 'x-tab-button')]",
        'dropdown_month': "//select[contains(@id, 'ddlBulan')]",
        'dropdown_year': "//select[contains(@id, 'ddlTahun')]",
        'dropdown_province': "//select[contains(@id, 'ddlProvinsi')]",
        'dropdown_city': "//select[contains(@id, 'ddlKota')]",
        'dropdown_bank': "//select[contains(@id, 'ddlBank')]",
        'checkbox_posisi_keuangan': "//input[contains(@id, 'chkLaporanPosisiKeuangan')]",
        'checkbox_laba_rugi': "//input[contains(@id, 'chkLaporanLabaRugi')]",
        'checkbox_kualitas_aset': "//input[contains(@id, 'chkLaporanKualitasAset')]",
        'button_tampilkan': "//input[contains(@value, 'Tampilkan') or contains(@id, 'btnTampilkan')]",
        'table_container': "//div[contains(@id, 'table') or contains(@class, 'table')]",
    }
    
    # Report type checkboxes (first 3 as per project plan)
    REPORT_TYPES = [
        'Laporan Posisi Keuangan',
        'Laporan Laba Rugi',
        'Laporan Kualitas Aset Produktif'
    ]


class Settings:
    """General application settings"""
    
    # Output settings
    OUTPUT_DIR = OUTPUT_DIR
    EXCEL_FILENAME_PREFIX = "ojk_report"
    EXCEL_FILENAME_DATE_FORMAT = "%Y%m%d_%H%M%S"
    
    # Logging settings
    LOGS_DIR = LOGS_DIR
    LOG_LEVEL = "INFO"
    LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    
    # Retry settings
    MAX_RETRIES = 3
    RETRY_DELAY = 5.0

