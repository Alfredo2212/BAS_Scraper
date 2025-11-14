"""
ASP.NET PostBack event handler
Handles dynamic element IDs and PostBack waiting logic
"""

import time
import random
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException

from config.settings import OJKConfig


class PostBackHandler:
    """Handles ASP.NET PostBack events and dynamic element interactions"""
    
    def __init__(self, driver: WebDriver, wait: WebDriverWait):
        """
        Initialize PostBack handler
        
        Args:
            driver: WebDriver instance
            wait: WebDriverWait instance
        """
        self.driver = driver
        self.wait = wait
    
    def wait_for_postback(self, old_element=None, new_element_selector: tuple = None, timeout: int = None):
        """
        Wait for PostBack to complete
        
        Args:
            old_element: Element that will become stale after PostBack (optional)
            new_element_selector: Tuple (By, selector) for element to wait for after PostBack
            timeout: Timeout in seconds. If None, uses OJKConfig.POSTBACK_WAIT_TIMEOUT
        """
        if timeout is None:
            timeout = OJKConfig.POSTBACK_WAIT_TIMEOUT
        
        # Wait for old element to become stale (if provided)
        if old_element:
            try:
                self.wait.until(EC.staleness_of(old_element))
            except (TimeoutException, StaleElementReferenceException):
                pass  # Element might already be stale or not found
        
        # Wait for new element to appear (if selector provided)
        if new_element_selector:
            self.wait.until(EC.presence_of_element_located(new_element_selector))
        
        # Additional delay to ensure PostBack is fully complete
        delay = OJKConfig.DELAY_AFTER_POSTBACK + random.uniform(*OJKConfig.RANDOM_DELAY_RANGE)
        time.sleep(delay)
    
    def select_dropdown_by_text(self, selector_key: str, value: str, wait_for_new_element: tuple = None):
        """
        Select value from dropdown and handle PostBack
        
        Args:
            selector_key: Key from OJKConfig.SELECTORS dictionary
            value: Text value to select
            wait_for_new_element: Optional tuple (By, selector) to wait for after PostBack
            
        Returns:
            The Select object for further operations if needed
        """
        # Get selector from config
        xpath = OJKConfig.SELECTORS.get(selector_key)
        if not xpath:
            raise ValueError(f"Selector key '{selector_key}' not found in OJKConfig.SELECTORS")
        
        # Find dropdown element
        dropdown_element = self.wait.until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        
        # Create Select object
        select = Select(dropdown_element)
        
        # Select by visible text
        select.select_by_visible_text(value)
        
        # Wait for PostBack
        self.wait_for_postback(
            old_element=dropdown_element,
            new_element_selector=wait_for_new_element
        )
        
        return select
    
    def select_dropdown_by_index(self, selector_key: str, index: int, wait_for_new_element: tuple = None):
        """
        Select value from dropdown by index and handle PostBack
        
        Args:
            selector_key: Key from OJKConfig.SELECTORS dictionary
            index: Index to select (0-based)
            wait_for_new_element: Optional tuple (By, selector) to wait for after PostBack
            
        Returns:
            The Select object for further operations if needed
        """
        # Get selector from config
        xpath = OJKConfig.SELECTORS.get(selector_key)
        if not xpath:
            raise ValueError(f"Selector key '{selector_key}' not found in OJKConfig.SELECTORS")
        
        # Find dropdown element
        dropdown_element = self.wait.until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        
        # Create Select object
        select = Select(dropdown_element)
        
        # Select by index
        select.select_by_index(index)
        
        # Wait for PostBack
        self.wait_for_postback(
            old_element=dropdown_element,
            new_element_selector=wait_for_new_element
        )
        
        return select
    
    def click_checkbox(self, selector_key: str):
        """
        Click a checkbox
        
        Args:
            selector_key: Key from OJKConfig.SELECTORS dictionary
        """
        xpath = OJKConfig.SELECTORS.get(selector_key)
        if not xpath:
            raise ValueError(f"Selector key '{selector_key}' not found in OJKConfig.SELECTORS")
        
        checkbox = self.wait.until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        
        # Check if not already selected
        if not checkbox.is_selected():
            checkbox.click()
            time.sleep(0.5)  # Small delay after clicking
    
    def click_button(self, selector_key: str):
        """
        Click a button and wait for result
        
        Args:
            selector_key: Key from OJKConfig.SELECTORS dictionary
        """
        xpath = OJKConfig.SELECTORS.get(selector_key)
        if not xpath:
            raise ValueError(f"Selector key '{selector_key}' not found in OJKConfig.SELECTORS")
        
        button = self.wait.until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        button.click()
        
        # Wait a bit for the action to start
        time.sleep(1)

