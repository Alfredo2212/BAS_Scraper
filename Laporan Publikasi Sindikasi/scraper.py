"""
Sindikasi Scraper
Finds listed BPRs from both BPR Konvensional and BPR Syariah report pages
"""

import time
import logging
import re
from pathlib import Path
from datetime import datetime
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

# Handle imports for both package and direct execution
try:
    from helper import ExtJSHelper
    from selenium_setup import SeleniumSetup
except ImportError:
    # If relative imports fail, try absolute imports
    import sys
    from pathlib import Path
    module_dir = Path(__file__).parent.parent / "Laporan Publikasi BPR Konvensional"
    if str(module_dir) not in sys.path:
        sys.path.insert(0, str(module_dir))
    from helper import ExtJSHelper
    from selenium_setup import SeleniumSetup

from config.settings import OJKConfig


class SindikasiScraper:
    """Scraper for finding BPRs in Sindikasi reports"""
    
    # URLs for both report types
    URL_KONVENSIONAL = "https://cfs.ojk.go.id/cfs/Report.aspx?BankTypeCode=BPK&BankTypeName=BPR%20Konvensional"
    URL_SYARIAH = "https://cfs.ojk.go.id/cfs/Report.aspx?BankTypeCode=BPS&BankTypeName=BPR%20Syariah"
    
    def __init__(self, headless: bool = False):
        """
        Initialize the scraper
        
        Args:
            headless: Whether to run browser in headless mode (default: False for visibility)
        """
        self.driver: WebDriver = None
        self.wait: WebDriverWait = None
        self.extjs: ExtJSHelper = None
        self.headless = headless
        
        # Setup logging
        log_dir = Path(__file__).parent.parent / "logs"
        log_dir.mkdir(exist_ok=True)
        log_file = log_dir / f"sindikasi_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
        self.logger.info(f"Log file: {log_file}")
    
    def initialize(self):
        """Initialize WebDriver and ExtJS helper"""
        if self.driver is None:
            self.driver = SeleniumSetup.create_driver(headless=self.headless)
            self.wait = SeleniumSetup.create_wait(self.driver)
            self.extjs = ExtJSHelper(self.driver, self.wait)
            
            # Don't minimize - keep window visible for monitoring
            self.logger.info("Chrome browser initialized")
    
    def cleanup(self, kill_processes: bool = True):
        """Cleanup resources"""
        if self.driver:
            try:
                self.driver.quit()
                self.logger.info("Browser closed")
            except Exception as e:
                self.logger.error(f"Error closing browser: {e}")
        
        if kill_processes:
            try:
                import sys
                from pathlib import Path
                module_dir = Path(__file__).parent.parent / "Laporan Publikasi BPR Konvensional"
                if str(module_dir) not in sys.path:
                    sys.path.insert(0, str(module_dir))
                from utils import kill_chrome_processes
                kill_chrome_processes()
            except Exception as e:
                self.logger.debug(f"Error killing Chrome processes: {e}")
    
    def read_list_file(self, file_path: Path) -> dict:
        """
        Read and parse list file
        
        Args:
            file_path: Path to the list file
            
        Returns:
            dict with keys: 'scrape' (bool), 'name' (str), 'banks' (list)
        """
        result = {
            'scrape': False,
            'name': '',
            'banks': []
        }
        
        if not file_path.exists():
            self.logger.error(f"List file not found: {file_path}")
            return result
        
        self.logger.info(f"Reading list file: {file_path}")
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            
            for line in lines:
                line = line.strip()
                
                # Skip empty lines
                if not line:
                    continue
                
                # Parse SCRAPE flag
                if line.upper().startswith('SCRAPE'):
                    scrape_match = re.search(r'SCRAPE\s*=\s*(TRUE|FALSE)', line, re.IGNORECASE)
                    if scrape_match:
                        result['scrape'] = scrape_match.group(1).upper() == 'TRUE'
                
                # Parse NAME
                elif line.upper().startswith('NAME'):
                    name_match = re.search(r'NAME\s*=\s*(.+)', line, re.IGNORECASE)
                    if name_match:
                        result['name'] = name_match.group(1).strip()
                
                # Bank names (everything else after NAME)
                elif result['name']:  # Only add banks after NAME is found
                    result['banks'].append(line)
            
            self.logger.info(f"SCRAPE = {result['scrape']}, NAME = {result['name']}")
            self.logger.info(f"Found {len(result['banks'])} BPRs in list")
            
            return result
            
        except Exception as e:
            self.logger.error(f"Error reading list file: {e}")
            return result
    
    def navigate_to_page(self, url: str):
        """Navigate to a report page and wait for ExtJS to load"""
        if self.driver is None:
            self.initialize()
        
        # Switch to default content in case we're in an iframe
        try:
            self.driver.switch_to.default_content()
        except:
            pass
        
        self.logger.info(f"Navigating to: {url}")
        self.driver.get(url)
        
        # Wait for page to load
        time.sleep(0.75)
        
        # Check if page is in iframe or main page
        max_attempts = 10
        for attempt in range(max_attempts):
            try:
                if self.extjs.check_extjs_available():
                    self.logger.debug("ExtJS is available in main page")
                    return
            except:
                pass
            time.sleep(0.75)
        
        # If not in main page, check for iframes
        iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
        if iframes:
            self.logger.debug(f"Found {len(iframes)} iframe(s), checking inside...")
            for i, iframe in enumerate(iframes):
                try:
                    self.driver.switch_to.frame(iframe)
                    time.sleep(1.125)
                    
                    if self.extjs.check_extjs_available():
                        self.logger.debug(f"ExtJS is available in iframe {i+1}")
                        return
                    
                    self.driver.switch_to.default_content()
                except:
                    self.driver.switch_to.default_content()
                    continue
        
        # Final check
        time.sleep(0.75)
        if self.extjs.check_extjs_available():
            self.logger.debug("ExtJS is now available")
            return
        
        self.logger.warning("ExtJS not immediately available, continuing anyway...")
    
    def _select_month(self, month: str):
        """Select month in the dropdown"""
        try:
            # Click month dropdown trigger
            trigger = self.driver.find_element(By.ID, "ext-gen1050")
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", trigger)
            time.sleep(0.75)
            self.driver.execute_script("arguments[0].click();", trigger)
            time.sleep(0.75)
            
            # Wait for dropdown and find month
            wait = WebDriverWait(self.driver, 5)
            wait.until(EC.presence_of_element_located((By.XPATH, "//ul[contains(@class, 'x-list-plain')]")))
            
            li_elements = self.driver.find_elements(By.XPATH, "//li[@role='option' or contains(@class, 'x-boundlist-item')]")
            if not li_elements:
                li_elements = self.driver.find_elements(By.XPATH, "//ul[contains(@class, 'x-list-plain')]//li")
            
            for li in li_elements:
                if li.text.strip().lower() == month.lower():
                    self.driver.execute_script("arguments[0].click();", li)
                    time.sleep(1.125)
                    self.logger.info(f"  [OK] Selected month: {month}")
                    return
        except Exception as e:
            self.logger.warning(f"  [WARNING] Error selecting month {month}: {e}")
    
    def _select_year(self, year: str):
        """Select year in the input field"""
        try:
            year_input = self.driver.find_element(By.ID, "Year-inputEl")
            year_input.clear()
            year_input.send_keys(year)
            year_input.send_keys(Keys.TAB)
            time.sleep(0.75)
            self.logger.info(f"  [OK] Selected year: {year}")
        except Exception as e:
            self.logger.warning(f"  [WARNING] Error selecting year {year}: {e}")
    
    def _search_bank_by_name(self, search_text: str, url_type: str = "") -> list:
        """
        Search for banks by typing in the search field
        Uses the BankCodeSearchField-inputEl input field
        
        Args:
            search_text: Bank name to search (without BPR/BPRS prefix)
            url_type: Type of URL (e.g., "BPR Konvensional" or "BPR Syariah") for filtering
            
        Returns:
            List of matching bank names from search results
        """
        try:
            # Step 1: Click dropdown trigger to open the bank dropdown
            self.logger.debug("  Clicking bank dropdown trigger (ext-gen1069)...")
            dropdown_trigger = self.driver.find_element(By.ID, "ext-gen1069")
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown_trigger)
            time.sleep(0.75)
            self.driver.execute_script("arguments[0].click();", dropdown_trigger)
            time.sleep(0.75)
            
            # Step 2: Wait for dropdown to appear and find search input
            wait = WebDriverWait(self.driver, 10)
            wait.until(EC.presence_of_element_located((By.ID, "BankCodeSearchField-inputEl")))
            search_input = self.driver.find_element(By.ID, "BankCodeSearchField-inputEl")
            
            # Step 3: Retry logic - wait after clicking dropdown, then type and check results
            self.logger.debug("  Waiting for dropdown to fully load, then searching...")
            bank_names = []
            max_retries = 3
            wait_time = 2.0
            
            for attempt in range(1, max_retries + 1):
                # Wait 2 seconds after clicking dropdown (or before retyping)
                self.logger.debug(f"  Attempt {attempt}/{max_retries}: Waiting {wait_time} seconds for dropdown to be ready...")
                time.sleep(wait_time)
                
                # Type (or retype) the search text
                self.logger.debug(f"  Typing search text: '{search_text}'")
                search_input.clear()
                search_input.send_keys(search_text)
                
                # Wait 2 seconds for search results to appear
                self.logger.debug(f"  Waiting {wait_time} seconds for search results...")
                time.sleep(wait_time)
                
                # Try to get results
                try:
                    # Get all spans with class="x-tree-node-text"
                    all_spans = self.driver.find_elements(By.XPATH, "//span[@class='x-tree-node-text' or contains(@class, 'x-tree-node-text')]")
                    
                    # Filter to only visible, non-empty spans
                    span_elements = [
                        s for s in all_spans
                        if s.is_displayed() and s.text.strip()
                    ]
                    
                    if span_elements:
                        skip_labels = ["Laporan Posisi Keuangan", "Laporan Laba Rugi", "Laporan", "Posisi", "Keuangan", "Laba", "Rugi"]
                        
                        # Add URL-type-specific labels to skip
                        if "konvensional" in url_type.lower():
                            skip_labels.append("BPK-BPR Konvensional")
                        if "syariah" in url_type.lower():
                            skip_labels.append("Komitmen dan Kontijensi")
                        
                        for span in span_elements:
                            span_text = span.text.strip()
                            
                            # Skip if it's a known label
                            is_label = any(label.lower() in span_text.lower() for label in skip_labels)
                            if is_label:
                                self.logger.debug(f"    Skipping label: {span_text[:60]}...")
                                continue
                            
                            # Bank names typically contain numbers (bank codes) or are reasonably long
                            has_number = any(char.isdigit() for char in span_text)
                            
                            # Only add if it looks like a bank name
                            if has_number or len(span_text) > 15:
                                bank_names.append(span_text)
                                self.logger.debug(f"    Found: {span_text[:60]}...")
                        
                        # If we found results, break out of retry loop
                        if bank_names:
                            self.logger.debug(f"  Found {len(bank_names)} result(s) on attempt {attempt}")
                            break
                        else:
                            self.logger.debug(f"  Attempt {attempt}: Found spans but no valid bank names yet")
                    else:
                        self.logger.debug(f"  Attempt {attempt}: No spans found yet")
                        
                except Exception as e:
                    self.logger.debug(f"  Attempt {attempt}: Error getting results: {e}")
            
            # Close dropdown
            try:
                search_input.send_keys(Keys.ESCAPE)
                time.sleep(0.5)
            except:
                pass
            
            if bank_names:
                self.logger.info(f"  Found {len(bank_names)} result(s) for search '{search_text}'")
            else:
                self.logger.info(f"  No results found for search '{search_text}' after {max_retries} attempts")
            
            return bank_names
            
        except Exception as e:
            self.logger.error(f"  [ERROR] Error searching for bank '{search_text}': {e}")
            import traceback
            self.logger.debug(traceback.format_exc())
            return []
    
    def _check_report_checkboxes(self, url_type: str):
        """
        Check the appropriate checkboxes based on URL type
        
        Args:
            url_type: Type of URL (e.g., "BPR Konvensional" or "BPR Syariah")
        """
        try:
            self.logger.info(f"  Checking report checkboxes for {url_type}...")
            
            # Find all checkboxes with class x-tree-checkbox
            # Try with role='checkbox' first (for BPRS), then fall back to just class (for Konvensional)
            wait = WebDriverWait(self.driver, 10)
            
            # Try to find checkboxes with role='checkbox' and class 'x-tree-checkbox'
            all_checkboxes = self.driver.find_elements(By.XPATH, "//*[@role='checkbox' and contains(@class, 'x-tree-checkbox')]")
            
            # If not found, try just class 'x-tree-checkbox' (for Konvensional)
            if not all_checkboxes:
                wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(@class, 'x-tree-checkbox')]")))
                all_checkboxes = self.driver.find_elements(By.XPATH, "//*[contains(@class, 'x-tree-checkbox')]")
            
            if not all_checkboxes:
                self.logger.warning("  No checkboxes found with class 'x-tree-checkbox'")
                return
            
            self.logger.debug(f"  Found {len(all_checkboxes)} checkbox(es) with class 'x-tree-checkbox'")
            
            # Determine which checkboxes to check
            if "konvensional" in url_type.lower():
                # Konvensional: check first 3 checkboxes (indices 0, 1, 2)
                indices_to_check = [0, 1, 2]
                self.logger.info("  Checking first 3 checkboxes for Konvensional")
            elif "syariah" in url_type.lower():
                # Syariah: check checkboxes 1, 2, and 4 (indices 0, 1, 3)
                indices_to_check = [0, 1, 3]
                self.logger.info("  Checking checkboxes 1, 2, and 4 for Syariah")
            else:
                self.logger.warning(f"  Unknown URL type: {url_type}, skipping checkbox selection")
                return
            
            # Check the specified checkboxes
            for idx in indices_to_check:
                if idx >= len(all_checkboxes):
                    self.logger.warning(f"  Checkbox index {idx} (position {idx + 1}) not available (only {len(all_checkboxes)} checkboxes found)")
                    continue
                
                checkbox = all_checkboxes[idx]
                try:
                    # Check if checkbox is visible
                    if not checkbox.is_displayed():
                        self.logger.debug(f"  Checkbox {idx + 1} is not visible, skipping")
                        continue
                    
                    # Scroll to checkbox
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", checkbox)
                    time.sleep(0.5)
                    
                    # Check if already checked
                    aria_checked = checkbox.get_attribute("aria-checked")
                    checkbox_type = checkbox.get_attribute("type")
                    is_selected = checkbox.is_selected() if checkbox_type == "checkbox" else False
                    
                    if aria_checked == "true" or is_selected:
                        self.logger.debug(f"  Checkbox {idx + 1} already checked")
                    else:
                        # Click checkbox using JavaScript
                        self.driver.execute_script("arguments[0].click();", checkbox)
                        self.logger.info(f"  [OK] Checked checkbox {idx + 1}")
                        time.sleep(0.5)
                        
                except Exception as e:
                    self.logger.warning(f"  Could not check checkbox {idx + 1}: {e}")
            
            self.logger.info("  [OK] Completed checkbox selection")
            
        except Exception as e:
            self.logger.error(f"  [ERROR] Error checking checkboxes: {e}")
            import traceback
            self.logger.debug(traceback.format_exc())
    
    def _select_bank_from_dropdown(self, bank_name: str) -> bool:
        """
        Select a bank from the dropdown by its name
        
        Args:
            bank_name: Full bank name to select (e.g., "620006-PT Bank Perekonomian Rakyat Syariah Hasanah Mandiri")
            
        Returns:
            True if successfully selected, False otherwise
        """
        try:
            self.logger.info(f"  Selecting bank: {bank_name[:80]}...")
            
            # Step 1: Click dropdown trigger to open the bank dropdown
            dropdown_trigger = self.driver.find_element(By.ID, "ext-gen1069")
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown_trigger)
            time.sleep(0.75)
            self.driver.execute_script("arguments[0].click();", dropdown_trigger)
            time.sleep(2.0)  # Wait for dropdown to fully open
            
            # Step 2: Find the bank in the dropdown by text
            wait = WebDriverWait(self.driver, 10)
            
            # Look for the bank name in spans with class x-tree-node-text
            bank_xpath = f"//span[@class='x-tree-node-text' or contains(@class, 'x-tree-node-text')][contains(text(), '{bank_name}')]"
            
            try:
                bank_span = wait.until(EC.element_to_be_clickable((By.XPATH, bank_xpath)))
                
                # Scroll to the bank option
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", bank_span)
                time.sleep(0.5)
                
                # Click the bank option
                self.driver.execute_script("arguments[0].click();", bank_span)
                time.sleep(1.0)  # Wait for selection to register
                
                self.logger.info(f"  [OK] Successfully selected bank: {bank_name[:80]}...")
                return True
                
            except Exception as e:
                self.logger.warning(f"  Could not find or click bank '{bank_name[:60]}...': {e}")
                # Try to close dropdown
                try:
                    dropdown_trigger.send_keys(Keys.ESCAPE)
                    time.sleep(0.5)
                except:
                    pass
                return False
                
        except Exception as e:
            self.logger.error(f"  [ERROR] Error selecting bank '{bank_name[:60]}...': {e}")
            import traceback
            self.logger.debug(traceback.format_exc())
            return False
    
    def _get_target_month_year(self) -> tuple[str, str]:
        """
        Determine target month and year based on current date.
        Uses quarterly mapping logic:
        - Jan, Feb, Mar → use December (YYYY-1)
        - Apr, May, Jun → use March YYYY
        - Jul, Aug, Sep → use June YYYY
        - Oct, Nov, Dec → use September YYYY
        Where YYYY is the current server year.
        
        Returns:
            Tuple of (month_name, year) e.g., ("September", "2025")
        """
        from datetime import datetime
        
        now = datetime.now()
        current_month = now.month
        current_year = now.year
        
        # Quarterly mapping logic
        month_names = {
            3: "Maret",      # 03
            6: "Juni",       # 06
            9: "September",  # 09
            12: "Desember"   # 12
        }
        
        # Map current month to quarterly report month
        if current_month in [1, 2, 3]:  # Jan, Feb, Mar
            target_month_num = 12  # Desember
            target_year = current_year - 1
        elif current_month in [4, 5, 6]:  # Apr, May, Jun
            target_month_num = 3  # Maret
            target_year = current_year
        elif current_month in [7, 8, 9]:  # Jul, Aug, Sep
            target_month_num = 6  # Juni
            target_year = current_year
        else:  # Oct, Nov, Dec
            target_month_num = 9  # September
            target_year = current_year
        
        month_name = month_names[target_month_num]
        
        self.logger.info(f"Current date: {now.strftime('%B %Y')}")
        self.logger.info(f"Target month/year: {month_name} {target_year}")
        
        return month_name, str(target_year)
    
    def _remove_bpr_prefix(self, bank_name: str) -> str:
        """
        Remove BPR/BPRS prefix from bank name for matching
        
        Args:
            bank_name: Bank name with or without prefix
            
        Returns:
            Bank name without BPR/BPRS prefix
        """
        # Remove common prefixes
        name = bank_name.strip()
        name = re.sub(r'^(BPRS|BPR|PT\s*BANK|BANK)\s+', '', name, flags=re.IGNORECASE)
        return name.strip()
    
    def _fuzzy_match_bank(self, search_name: str, bank_list: list) -> list:
        """
        Fuzzy match bank name against list (without BPR/BPRS prefix)
        Returns all matching banks
        
        Args:
            search_name: Bank name to search for (e.g., "SYARIKAT MADANI")
            bank_list: List of available bank names from dropdown
            
        Returns:
            List of matching bank names (can be multiple)
        """
        # Remove BPR/BPRS prefix from search name
        search_clean = self._remove_bpr_prefix(search_name)
        search_upper = search_clean.upper().strip()
        
        matches = []
        
        for bank in bank_list:
            bank_clean = self._remove_bpr_prefix(bank)
            bank_upper = bank_clean.upper().strip()
            
            # Skip if either is empty after cleaning
            if not search_upper or not bank_upper:
                continue
            
            # Exact match
            if search_upper == bank_upper:
                matches.append(bank)
                continue
            
            # Partial match (search name is contained in bank name)
            if search_upper in bank_upper:
                matches.append(bank)
                continue
            
            # Reverse partial match (bank name is contained in search name)
            if bank_upper in search_upper:
                matches.append(bank)
                continue
            
            # Word-by-word matching (for cases like "SYARIKAT MADANI" matching "SYARIKAT MADANI INDONESIA")
            search_words = search_upper.split()
            bank_words = bank_upper.split()
            
            # If all search words are in bank words, it's a match
            if len(search_words) > 0 and all(word in bank_words for word in search_words):
                matches.append(bank)
                continue
        
        return matches
    
    def find_bank_in_url(self, bank_name: str, url: str, url_type: str, target_month: str, target_year: str) -> list:
        """
        Find a bank in a specific URL using search field (can return multiple matches)
        
        Args:
            bank_name: Bank name to search for
            url: URL to check
            url_type: Type of URL (e.g., "BPR Konvensional" or "BPR Syariah")
            target_month: Target month to select
            target_year: Target year to select
            
        Returns:
            List of matching bank names (empty list if not found)
        """
        try:
            self.logger.info(f"  Checking {url_type}...")
            self.navigate_to_page(url)
            
            # Step 1: Select month and year first
            self.logger.info(f"  Setting month: {target_month}, year: {target_year}")
            self._select_month(target_month)
            self._select_year(target_year)
            time.sleep(1.0)  # Wait for page to update after month/year selection
            
            # Step 2: Check the appropriate checkboxes
            self._check_report_checkboxes(url_type)
            time.sleep(1.0)  # Wait after checking checkboxes
            
            # Step 3: Remove BPR/BPRS prefix from search text
            search_text = self._remove_bpr_prefix(bank_name)
            self.logger.info(f"  Searching for: '{search_text}' (from '{bank_name}')")
            
            # Step 4: Search using the search field
            matches = self._search_bank_by_name(search_text, url_type)
            
            if matches:
                self.logger.info(f"  [OK] Found {len(matches)} match(es) in {url_type}:")
                for match in matches:
                    self.logger.info(f"    - {match}")
            else:
                self.logger.info(f"  Not found in {url_type}")
            
            return matches
            
        except Exception as e:
            self.logger.error(f"  [ERROR] Error checking {url_type}: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            return []
    
    def process_bank_matches(self, matches: list, bank_name: str, url_type: str):
        """
        Process each bank match by selecting it from the dropdown
        This will iterate through all matches and select each one
        
        Args:
            matches: List of matching bank names from search
            bank_name: Original bank name being searched
            url_type: Type of URL (e.g., "BPR Konvensional" or "BPR Syariah")
        """
        if not matches:
            self.logger.info(f"  No matches to process for '{bank_name}'")
            return
        
        self.logger.info(f"  Processing {len(matches)} match(es) for '{bank_name}'...")
        
        for idx, match in enumerate(matches, 1):
            self.logger.info(f"  [{idx}/{len(matches)}] Processing match: {match[:80]}...")
            
            # Select the bank from dropdown
            success = self._select_bank_from_dropdown(match)
            
            if success:
                self.logger.info(f"  [OK] Successfully selected bank {idx}/{len(matches)}")
                # TODO: Extract data for this bank (will be implemented next)
                time.sleep(1.0)  # Wait before processing next match
            else:
                self.logger.warning(f"  [WARNING] Failed to select bank {idx}/{len(matches)}")
            
            # Small delay between matches
            time.sleep(0.5)
    
    def _determine_bank_type(self, bank_name: str) -> str:
        """
        Determine which URL to check based on bank name
        
        Args:
            bank_name: Bank name to check
            
        Returns:
            'syariah' if bank contains BPRS or Syariah, 'konvensional' otherwise
        """
        bank_upper = bank_name.upper()
        
        # Check for BPRS or Syariah indicators
        if 'BPRS' in bank_upper or 'SYARIAH' in bank_upper:
            return 'syariah'
        else:
            return 'konvensional'
    
    def find_all_banks(self, list_file_path: Path):
        """
        Main orchestrator: Find all banks from list in appropriate URL
        
        Args:
            list_file_path: Path to the list file
        """
        # Read list file
        list_data = self.read_list_file(list_file_path)
        
        if not list_data['scrape']:
            self.logger.info("SCRAPE = FALSE, skipping bank search")
            return
        
        if not list_data['banks']:
            self.logger.warning("No banks found in list file")
            return
        
        # Get target month and year
        target_month, target_year = self._get_target_month_year()
        self.logger.info(f"Target period: {target_month} {target_year}")
        
        self.logger.info("=" * 70)
        self.logger.info("Starting bank search...")
        self.logger.info("=" * 70)
        
        # Initialize browser
        self.initialize()
        
        # Statistics
        found_konvensional = []  # List of (search_name, [matched_banks])
        found_syariah = []  # List of (search_name, [matched_banks])
        not_found = []
        
        try:
            # Process each bank
            for i, bank_name in enumerate(list_data['banks'], 1):
                self.logger.info("")
                self.logger.info("=" * 70)
                self.logger.info(f"[{i}/{len(list_data['banks'])}] Searching for: {bank_name}")
                self.logger.info("=" * 70)
                
                # Determine which URL to check based on bank name
                bank_type = self._determine_bank_type(bank_name)
                
                if bank_type == 'syariah':
                    self.logger.info(f"  Bank type: BPR Syariah (contains BPRS or Syariah)")
                    matches = self.find_bank_in_url(
                        bank_name,
                        self.URL_SYARIAH,
                        "BPR Syariah",
                        target_month,
                        target_year
                    )
                    if matches:
                        found_syariah.append((bank_name, matches))
                        self.logger.info(f"  [SUCCESS] Found {len(matches)} match(es) for '{bank_name}' in BPR Syariah")
                        # Process each match by selecting it from dropdown
                        self.process_bank_matches(matches, bank_name, "BPR Syariah")
                    else:
                        not_found.append(bank_name)
                        self.logger.warning(f"  [WARNING] Not found in BPR Syariah")
                else:
                    self.logger.info(f"  Bank type: BPR Konvensional")
                    matches = self.find_bank_in_url(
                        bank_name,
                        self.URL_KONVENSIONAL,
                        "BPR Konvensional",
                        target_month,
                        target_year
                    )
                    if matches:
                        found_konvensional.append((bank_name, matches))
                        self.logger.info(f"  [SUCCESS] Found {len(matches)} match(es) for '{bank_name}' in BPR Konvensional")
                        # Process each match by selecting it from dropdown
                        self.process_bank_matches(matches, bank_name, "BPR Konvensional")
                    else:
                        not_found.append(bank_name)
                        self.logger.warning(f"  [WARNING] Not found in BPR Konvensional")
            
            # Print summary
            self.logger.info("")
            self.logger.info("=" * 70)
            self.logger.info("SEARCH SUMMARY")
            self.logger.info("=" * 70)
            self.logger.info(f"Target period: {target_month} {target_year}")
            self.logger.info(f"Total banks searched: {len(list_data['banks'])}")
            self.logger.info(f"Found in BPR Konvensional: {len(found_konvensional)}")
            self.logger.info(f"Found in BPR Syariah: {len(found_syariah)}")
            self.logger.info(f"Not found: {len(not_found)}")
            
            if found_konvensional:
                self.logger.info("")
                self.logger.info("Banks found in BPR Konvensional:")
                for search_name, matches in found_konvensional:
                    self.logger.info(f"  Search: '{search_name}' → Found {len(matches)} match(es):")
                    for match in matches:
                        self.logger.info(f"    - {match}")
            
            if found_syariah:
                self.logger.info("")
                self.logger.info("Banks found in BPR Syariah:")
                for search_name, matches in found_syariah:
                    self.logger.info(f"  Search: '{search_name}' → Found {len(matches)} match(es):")
                    for match in matches:
                        self.logger.info(f"    - {match}")
            
            if not_found:
                self.logger.info("")
                self.logger.warning("Banks NOT found in their respective URLs:")
                for bank in not_found:
                    self.logger.warning(f"  - {bank}")
            
            self.logger.info("=" * 70)
            
        except Exception as e:
            self.logger.error(f"Error during bank search: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
        
        finally:
            self.cleanup()

