"""
Selenium WebDriver setup and configuration
Handles browser initialization with proper settings
"""

import random
import os
import shutil
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.core.os_manager import ChromeType

from config.settings import OJKConfig


class SeleniumSetup:
    """Handles Selenium WebDriver setup and configuration"""
    
    @staticmethod
    def create_driver(headless: bool = None) -> webdriver.Chrome:
        """
        Create and configure Chrome WebDriver
        
        Args:
            headless: Whether to run in headless mode. If None, uses OJKConfig.HEADLESS_MODE
            
        Returns:
            Configured Chrome WebDriver instance
        """
        if headless is None:
            headless = OJKConfig.HEADLESS_MODE
        
        # Chrome options
        chrome_options = Options()
        
        if headless:
            chrome_options.add_argument('--headless')
        
        # Window size
        chrome_options.add_argument(f'--window-size={OJKConfig.WINDOW_SIZE[0]},{OJKConfig.WINDOW_SIZE[1]}')
        
        # Additional options for stability
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # Random user agent
        user_agent = random.choice(OJKConfig.USER_AGENTS)
        chrome_options.add_argument(f'user-agent={user_agent}')
        
        # Create service with automatic driver management
        # Try to get driver path, with fallback for cache issues
        try:
            driver_path = ChromeDriverManager().install()
            # Fix: ChromeDriverManager sometimes returns wrong file (e.g., THIRD_PARTY_NOTICES)
            # Check if it's actually the executable, if not, find chromedriver.exe in the same directory
            driver_path_obj = Path(driver_path)
            if not driver_path_obj.name.endswith('.exe'):
                # Look for chromedriver.exe in the same directory
                driver_dir = driver_path_obj.parent
                chromedriver_exe = driver_dir / "chromedriver.exe"
                if chromedriver_exe.exists():
                    driver_path = str(chromedriver_exe)
                else:
                    # Try parent directory
                    parent_dir = driver_dir.parent
                    chromedriver_exe = parent_dir / "chromedriver.exe"
                    if chromedriver_exe.exists():
                        driver_path = str(chromedriver_exe)
        except Exception as e:
            print(f"[WARNING] ChromeDriverManager failed: {e}")
            print("[INFO] Attempting to clear cache and retry...")
            # Clear cache and retry
            cache_path = Path.home() / ".wdm"
            if cache_path.exists():
                try:
                    shutil.rmtree(cache_path)
                    print("[INFO] Cache cleared, retrying...")
                except Exception as clear_error:
                    print(f"[WARNING] Could not clear cache: {clear_error}")
            driver_path = ChromeDriverManager().install()
            # Apply the same fix after retry
            driver_path_obj = Path(driver_path)
            if not driver_path_obj.name.endswith('.exe'):
                driver_dir = driver_path_obj.parent
                chromedriver_exe = driver_dir / "chromedriver.exe"
                if chromedriver_exe.exists():
                    driver_path = str(chromedriver_exe)
        
        service = Service(driver_path)
        
        # Create driver
        driver = webdriver.Chrome(service=service, options=chrome_options)
        
        # Set timeouts
        driver.set_page_load_timeout(OJKConfig.PAGE_LOAD_TIMEOUT)
        driver.implicitly_wait(5)  # Implicit wait for element finding
        
        return driver
    
    @staticmethod
    def create_wait(driver: webdriver.Chrome, timeout: int = None) -> WebDriverWait:
        """
        Create WebDriverWait instance
        
        Args:
            driver: WebDriver instance
            timeout: Wait timeout in seconds. If None, uses OJKConfig.ELEMENT_WAIT_TIMEOUT
            
        Returns:
            WebDriverWait instance
        """
        if timeout is None:
            timeout = OJKConfig.ELEMENT_WAIT_TIMEOUT
        
        return WebDriverWait(driver, timeout)

