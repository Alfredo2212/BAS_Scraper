"""
Data extraction from HTML tables
Parses table data using BeautifulSoup
"""

from bs4 import BeautifulSoup
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

from config.settings import OJKConfig


class DataExtractor:
    """Handles data extraction from HTML tables"""
    
    def __init__(self, driver: WebDriver, wait: WebDriverWait):
        """
        Initialize data extractor
        
        Args:
            driver: WebDriver instance
            wait: WebDriverWait instance
        """
        self.driver = driver
        self.wait = wait
    
    def extract_table_data(self, timeout: int = None) -> dict:
        """
        Extract data from the results table
        
        Args:
            timeout: Timeout for waiting for table. If None, uses OJKConfig.ELEMENT_WAIT_TIMEOUT
            
        Returns:
            Dictionary with 'success' flag and 'data' (list of dicts) or 'message' (error message)
        """
        if timeout is None:
            timeout = OJKConfig.ELEMENT_WAIT_TIMEOUT
        
        try:
            # Wait for table container to appear
            table_container = self.wait.until(
                EC.presence_of_element_located((By.XPATH, OJKConfig.SELECTORS['table_container']))
            )
            
            # Get innerHTML of the table container
            table_html = table_container.get_attribute('innerHTML')
            
            if not table_html or not table_html.strip():
                return {
                    "success": False,
                    "message": "Table container found but is empty"
                }
            
            # Parse HTML with BeautifulSoup
            soup = BeautifulSoup(table_html, 'html.parser')
            
            # Find all table rows
            rows = soup.find_all('tr')
            
            if not rows:
                return {
                    "success": False,
                    "message": "No table rows found"
                }
            
            # Extract data from rows
            data = []
            headers = []
            
            for i, row in enumerate(rows):
                cells = row.find_all(['td', 'th'])
                cell_texts = [cell.get_text(strip=True) for cell in cells]
                
                if not cell_texts:
                    continue
                
                # First row might be headers
                if i == 0 and row.find('th'):
                    headers = cell_texts
                    continue
                
                # Create dictionary from row data
                if headers:
                    row_dict = dict(zip(headers, cell_texts))
                else:
                    # If no headers, use column indices
                    row_dict = {f"Column_{j+1}": text for j, text in enumerate(cell_texts)}
                
                data.append(row_dict)
            
            return {
                "success": True,
                "data": data,
                "headers": headers if headers else None
            }
            
        except TimeoutException:
            return {
                "success": False,
                "message": "Table container not found - no data available for this combination"
            }
        except Exception as e:
            return {
                "success": False,
                "message": f"Error extracting table data: {str(e)}"
            }

