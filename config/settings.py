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
    
    # Timeouts (in seconds) - reduced by 50%
    PAGE_LOAD_TIMEOUT = 15  # Reduced from 30
    ELEMENT_WAIT_TIMEOUT = 15  # Reduced from 30
    POSTBACK_WAIT_TIMEOUT = 15  # Reduced from 30
    
    # Delays (in seconds) - to avoid rate limiting (reduced by 50%)
    DELAY_BETWEEN_REQUESTS = 1.0  # Reduced from 2.0
    DELAY_AFTER_POSTBACK = 1.5  # Reduced from 3.0
    RANDOM_DELAY_RANGE = (0, 0.5)  # Reduced from (0, 1)
    
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
        # ExtJS combobox selectors (not standard select elements)
        'dropdown_month': "//input[@name='Month' or contains(@id, 'Month') or @id='Month-inputEl']",
        'dropdown_month_trigger': "//input[@name='Month' or contains(@id, 'Month')]/following-sibling::td[contains(@class, 'x-trigger-cell')]//div[contains(@class, 'x-form-trigger')]",
        'dropdown_year': "//input[@name='Year' or contains(@id, 'Year') or @id='Year-inputEl']",
        'dropdown_year_trigger': "//input[@name='Year' or contains(@id, 'Year')]/following-sibling::td[contains(@class, 'x-trigger-cell')]//div[contains(@class, 'x-form-trigger')]",
        'dropdown_province': "//input[@name='Province' or contains(@id, 'Province') or contains(@id, 'Provinsi')]",
        'dropdown_province_trigger': "//input[@name='Province' or contains(@id, 'Province') or contains(@id, 'Provinsi')]/following-sibling::td[contains(@class, 'x-trigger-cell')]//div[contains(@class, 'x-form-trigger')]",
        'dropdown_city': "//input[@name='City' or contains(@id, 'City') or contains(@id, 'Kota')]",
        'dropdown_city_trigger': "//input[@name='City' or contains(@id, 'City') or contains(@id, 'Kota')]/following-sibling::td[contains(@class, 'x-trigger-cell')]//div[contains(@class, 'x-form-trigger')]",
        'dropdown_bank': "//input[@name='Bank' or contains(@id, 'Bank')]",
        'dropdown_bank_trigger': "//input[@name='Bank' or contains(@id, 'Bank')]/following-sibling::td[contains(@class, 'x-trigger-cell')]//div[contains(@class, 'x-form-trigger')]",
        # ExtJS dropdown menu items (appears after clicking trigger)
        'dropdown_menu': "//div[contains(@class, 'x-boundlist')]",
        'dropdown_menu_list': "//ul[contains(@class, 'x-list-plain')]",
        'dropdown_menu_item': "//li[@role='option' or contains(@class, 'x-boundlist-item')]",
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

