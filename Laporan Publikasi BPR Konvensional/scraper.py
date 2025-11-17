"""
OJK ExtJS Scraper
Main scraping logic using ExtJS API exclusively
No DOM clicking, pure ExtJS ComponentQuery
"""

import time
import csv
import re
from pathlib import Path
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
except ImportError:
    print("[WARNING] openpyxl not installed. Excel export will not work. Install with: pip install openpyxl")
    Workbook = None

# Handle imports for both package and direct execution
try:
    from .helper import ExtJSHelper
    from .selenium_setup import SeleniumSetup
except ImportError:
    # If relative imports fail, try absolute imports
    import sys
    from pathlib import Path
    module_dir = Path(__file__).parent
    if str(module_dir) not in sys.path:
        sys.path.insert(0, str(module_dir))
    from helper import ExtJSHelper
    from selenium_setup import SeleniumSetup

from config.settings import OJKConfig, Settings


class OJKExtJSScraper:
    """Main scraper class using ExtJS API exclusively"""
    
    # Direct URL to BPR Konvensional report page
    REPORT_URL = "https://cfs.ojk.go.id/cfs/Report.aspx?BankTypeCode=BPK&BankTypeName=BPR%20Konvensional"
    
    def __init__(self, headless: bool = None):
        """
        Initialize the scraper
        
        Args:
            headless: Whether to run browser in headless mode
        """
        self.base_url = self.REPORT_URL  # Use direct report URL
        self.driver: WebDriver = None
        self.wait: WebDriverWait = None
        self.extjs: ExtJSHelper = None
        self.headless = headless if headless is not None else OJKConfig.HEADLESS_MODE
        self.output_dir = Path(Settings.OUTPUT_DIR)
        self.output_dir.mkdir(exist_ok=True)
        self.excel_wb = None  # Workbook for appending data
        self.excel_ws = None  # Worksheet for appending data
        self.excel_row = 1  # Current row in Excel
    
    def initialize(self):
        """Initialize WebDriver and ExtJS helper"""
        if self.driver is None:
            self.driver = SeleniumSetup.create_driver(headless=self.headless)
            self.wait = SeleniumSetup.create_wait(self.driver)
            self.extjs = ExtJSHelper(self.driver, self.wait)
    
    def navigate_to_page(self):
        """Navigate directly to BPR Konvensional report page"""
        if self.driver is None:
            self.initialize()
        
        print(f"[INFO] Navigating directly to report page: {self.base_url}")
        self.driver.get(self.base_url)
        
        # Wait for page to load
        print("[INFO] Waiting for page to fully load...")
        time.sleep(0.5)
        
        # Check if page is in iframe or main page
        # Try to find ExtJS in main page first
        print("[INFO] Checking for ExtJS in main page...")
        max_attempts = 10
        for attempt in range(max_attempts):
            try:
                if self.extjs.check_extjs_available():
                    print("[OK] ExtJS is available in main page")
                    return
            except:
                pass
            time.sleep(0.5)
        
        # If not in main page, check for iframes
        print("[INFO] ExtJS not in main page, checking for iframes...")
        iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
        if iframes:
            print(f"[INFO] Found {len(iframes)} iframe(s), checking inside...")
            for i, iframe in enumerate(iframes):
                try:
                    self.driver.switch_to.frame(iframe)
                    print(f"[INFO] Switched to iframe {i+1}")
                    time.sleep(0.5)  # Reduced from 2s to 0.4s (80% faster)
                    
                    # Check for ExtJS in this iframe
                    if self.extjs.check_extjs_available():
                        print(f"[OK] ExtJS is available in iframe {i+1}")
                        return
                    
                    # Switch back to try next iframe
                    self.driver.switch_to.default_content()
                except:
                    self.driver.switch_to.default_content()
                    continue
        
        # Wait a bit more and check again (reduced by 80%)
        print("[INFO] Waiting for ExtJS to load...")
        time.sleep(0.5)
        
        # Final check
        if self.extjs.check_extjs_available():
            print("[OK] ExtJS is now available")
            return
        
        # Debug: Check JavaScript context
        print("[WARNING] ExtJS not immediately available, checking JavaScript context...")
        try:
            debug_js = """
            (function() {
                try {
                    return {
                        hasExt: typeof Ext !== 'undefined',
                        hasComponentQuery: typeof Ext !== 'undefined' && typeof Ext.ComponentQuery !== 'undefined',
                        hasWindow: typeof window !== 'undefined',
                        documentReady: document.readyState,
                        hasJQuery: typeof jQuery !== 'undefined',
                        url: window.location.href
                    };
                } catch (e) {
                    return {error: e.toString()};
                }
            })();
            """
            debug_result = self.driver.execute_script(debug_js)
            print(f"[DEBUG] JavaScript context: {debug_result}")
        except Exception as debug_error:
            print(f"[DEBUG] Could not execute debug script: {debug_error}")
        
        print("[WARNING] ExtJS not available yet, but will continue (it may load after page fully loads)")
    
    def scrape_all_data(self, month: str = "Desember", year: str = "2024"):
        """
        Main scraping loop
        Iterates through all provinces, cities, and banks
        
        Args:
            month: Month to select (e.g., "Desember")
            year: Year to select (e.g., "2024")
        """
        if self.driver is None:
            self.initialize()
        
        # Wait for page to fully load (reduced to 30% of original)
        print("[INFO] Waiting for page to fully load...")
        time.sleep(0.5)  # Reduced to 30% (0.6s * 0.3 = 0.18s)
        
        # Try to find and click the trigger arrow with ID ext-gen1050 (static ID for month dropdown)
        print("[INFO] Looking for trigger arrow (id='ext-gen1050')...")
        trigger_found = False
        max_attempts = 10
        wait_interval = 0.5
        
        for attempt in range(max_attempts):
            try:
                # Try to find by ID
                trigger = self.driver.find_element(By.ID, "ext-gen1050")
                print(f"[OK] Found trigger arrow (attempt {attempt + 1})")
                
                # Click the trigger to open dropdown
                print("[INFO] Clicking trigger arrow to open month dropdown...")
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", trigger)
                time.sleep(0.5)  # Reduced from 0.25s to 0.05s (80% faster)
                self.driver.execute_script("arguments[0].click();", trigger)
                print("[OK] Trigger arrow clicked")
                trigger_found = True
                break
            except:
                if attempt < max_attempts - 1:
                    print(f"[DEBUG] Trigger not found, waiting {wait_interval} seconds...")
                    time.sleep(wait_interval)
                else:
                    # Try alternative: find by class
                    try:
                        triggers = self.driver.find_elements(By.XPATH, "//div[contains(@class, 'x-form-arrow-trigger') and contains(@class, 'x-trigger-index-0')]")
                        if triggers:
                            trigger = triggers[0]
                            print("[OK] Found trigger arrow by class")
                            self.driver.execute_script("arguments[0].click();", trigger)
                            print("[OK] Trigger arrow clicked")
                            trigger_found = True
                            break
                    except:
                        pass
        
        if not trigger_found:
            print("[WARNING] Could not find trigger arrow, but will try ExtJS API method...")
        
        # Wait for dropdown to appear and find <li> element
        if trigger_found:
            print("[INFO] Waiting for dropdown menu to appear...")
            time.sleep(0.5)  # Reduced from 1s to 0.2s (80% faster)
            
            # Find and click the <li> element with the month value
            print(f"[INFO] Looking for <li> element with text '{month}'...")
            try:
                # Wait for dropdown list to appear
                from selenium.webdriver.support.ui import WebDriverWait
                from selenium.webdriver.support import expected_conditions as EC
                
                wait = WebDriverWait(self.driver, 5)
                dropdown_list = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//ul[contains(@class, 'x-list-plain')]"))
                )
                print("[OK] Dropdown menu appeared")
                
                # Find all <li> elements
                li_elements = self.driver.find_elements(By.XPATH, "//li[@role='option' or contains(@class, 'x-boundlist-item')]")
                if not li_elements:
                    li_elements = self.driver.find_elements(By.XPATH, "//ul[contains(@class, 'x-list-plain')]//li")
                
                print(f"[DEBUG] Found {len(li_elements)} <li> elements in dropdown")
                
                # Find and click the matching <li>
                target_li = None
                for li in li_elements:
                    try:
                        li_text = li.text.strip()
                        print(f"[DEBUG] Checking <li>: '{li_text}'")
                        if li_text.lower() == month.lower():
                            target_li = li
                            print(f"[OK] Found matching <li> element: '{li_text}'")
                            break
                    except:
                        continue
                
                if target_li:
                    print(f"[INFO] Clicking <li> element with text '{month}'...")
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target_li)
                    time.sleep(0.5)
                    self.driver.execute_script("arguments[0].click();", target_li)
                    print(f"[OK] Clicked <li> element with text '{month}'")
                    time.sleep(0.5)  # Wait for PostBack
                    print("[INFO] Waiting 1 second buffer after month selection...")
                    time.sleep(0.5)  # Buffer between dropdown selections
                else:
                    available_options = [li.text.strip() for li in li_elements if li.text.strip()]
                    print(f"[WARNING] Could not find <li> with text '{month}'. Available: {available_options}")
            except Exception as e:
                print(f"[WARNING] Could not click <li> element: {e}")
        
        # Now try ExtJS API method for other operations
        print("[INFO] Checking for ExtJS availability...")
        max_attempts = 10
        for attempt in range(max_attempts):
            if self.extjs.check_extjs_available():
                print("[OK] ExtJS is available")
                break
            time.sleep(0.5)  # Reduced from 1s to 0.2s (80% faster)
        else:
            print("[WARNING] ExtJS not available, but will try to continue...")
        
        # List all combos for debugging
        print("[DEBUG] Listing all available comboboxes...")
        combos = self.extjs.list_all_combos()
        if combos:
            print(f"[OK] Found {len(combos)} comboboxes:")
            for combo in combos:
                print(f"  - Index {combo['index']}: name='{combo['name']}', id='{combo['id']}'")
        else:
            print("[WARNING] No comboboxes found yet")
            time.sleep(0.5)  # Wait a bit more
            combos = self.extjs.list_all_combos()
            if combos:
                print(f"[OK] Found {len(combos)} comboboxes after additional wait")
            else:
                print("[WARNING] No comboboxes found - will try to continue anyway")
        
        # Step 2: Select year by typing directly into input field
        print(f"\n[Step 2] Selecting year: 2024 (hardcoded)")
        time.sleep(0.5)
        
        try:
            # Find year input field by ID
            year_input = self.driver.find_element(By.ID, "Year-inputEl")
            print("[OK] Found year input field (Year-inputEl)")
            
            # Clear and type the year
            year_input.clear()
            year_input.send_keys("2024")
            print("[OK] Typed '2024' into year input field")
            
            # Trigger change event (press Tab)
            from selenium.webdriver.common.keys import Keys
            year_input.send_keys(Keys.TAB)
            time.sleep(0.5)  # Wait for PostBack
            print("[INFO] Waiting 1 second buffer after year selection...")
            time.sleep(0.5)  # Buffer between dropdown selections
        except Exception as e:
            print(f"[WARNING] Could not find year input field: {e}")
            # Fallback: try ExtJS API if available
            if combos:
                year_combo_name = self._find_combo_name_by_keyword("year")
                if year_combo_name:
                    self.extjs.set_extjs_combo(year_combo_name, "2024")
                else:
                    year_combo_name = self.extjs.find_combo_by_position(1)
                    if year_combo_name:
                        self.extjs.set_extjs_combo(year_combo_name, "2024")
        
        # Step 3: Select province by clicking dropdown arrow and <li> element
        print("\n[Step 3] Selecting province...")
        time.sleep(0.5)
        
        # Try to find and click the trigger arrow with ID ext-gen1059 (static ID for province dropdown)
        print("[INFO] Looking for province trigger arrow (id='ext-gen1059')...")
        province_trigger_found = False
        max_attempts = 10
        wait_interval = 0.5
        
        for attempt in range(max_attempts):
            try:
                # Try to find by ID
                province_trigger = self.driver.find_element(By.ID, "ext-gen1059")
                print(f"[OK] Found province trigger arrow (attempt {attempt + 1})")
                
                # Click the trigger to open dropdown
                print("[INFO] Clicking province trigger arrow to open dropdown...")
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", province_trigger)
                time.sleep(0.5)
                self.driver.execute_script("arguments[0].click();", province_trigger)
                print("[OK] Province trigger arrow clicked")
                province_trigger_found = True
                break
            except:
                if attempt < max_attempts - 1:
                    print(f"[DEBUG] Province trigger not found, waiting {wait_interval} seconds...")
                    time.sleep(wait_interval)
                else:
                    print("[WARNING] Could not find province trigger arrow")
        
        # Wait for dropdown to appear and find <li> element
        if province_trigger_found:
            print("[INFO] Waiting for province dropdown menu to appear...")
            time.sleep(0.5)
            
            # Find and click the <li> element with the province value
            province_name = "Provinsi Kep. Riau"
            print(f"[INFO] Looking for <li> element with text '{province_name}'...")
            try:
                # Wait for dropdown list to appear
                from selenium.webdriver.support.ui import WebDriverWait
                from selenium.webdriver.support import expected_conditions as EC
                
                wait = WebDriverWait(self.driver, 5)
                dropdown_list = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//ul[contains(@class, 'x-list-plain')]"))
                )
                print("[OK] Province dropdown menu appeared")
                
                # Find all <li> elements
                li_elements = self.driver.find_elements(By.XPATH, "//li[@role='option' or contains(@class, 'x-boundlist-item')]")
                if not li_elements:
                    li_elements = self.driver.find_elements(By.XPATH, "//ul[contains(@class, 'x-list-plain')]//li")
                
                print(f"[DEBUG] Found {len(li_elements)} <li> elements in province dropdown")
                
                # Find and click the matching <li>
                target_li = None
                for li in li_elements:
                    try:
                        li_text = li.text.strip()
                        if not li_text:  # Skip empty elements
                            continue
                        print(f"[DEBUG] Checking <li>: '{li_text}'")
                        # Match if province name is contained in li_text or vice versa (but not empty)
                        if (province_name.lower() in li_text.lower() or li_text.lower() in province_name.lower()) and li_text:
                            target_li = li
                            print(f"[OK] Found matching <li> element: '{li_text}'")
                            break
                    except:
                        continue
                
                if target_li:
                    print(f"[INFO] Clicking <li> element with text '{province_name}'...")
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target_li)
                    time.sleep(0.5)
                    self.driver.execute_script("arguments[0].click();", target_li)
                    print(f"[OK] Clicked <li> element with text '{province_name}'")
                    time.sleep(0.5)  # Wait for PostBack
                    print("[INFO] Waiting 1 second buffer after province selection...")
                    time.sleep(0.5)  # Buffer between dropdown selections
                else:
                    available_options = [li.text.strip() for li in li_elements if li.text.strip()]
                    print(f"[WARNING] Could not find <li> with text '{province_name}'. Available: {available_options[:10]}...")
            except Exception as e:
                print(f"[WARNING] Could not click province <li> element: {e}")
        
        # Step 4: Select initial dropdowns and checkboxes (only once at start)
        print("\n[Step 4] Starting initial dropdown and checkbox selection...")
        self._select_initial_dropdowns_and_checkboxes()
        
        # Wait a bit after checkbox is ticked to ensure dropdowns are ready
        print("\n[INFO] Waiting for dropdowns to be ready after checkbox selection...")
        time.sleep(2.0)
        
        # Initialize Excel file
        self._initialize_excel(year)
        
        # Step 5: Sequential iteration through cities and banks
        # This starts AFTER checkbox is ticked - we iterate through all city/bank combinations
        print("\n[Step 5] Starting sequential iteration through all cities and banks...")
        print("[INFO] Iteration starts after checkbox is ticked - will select city, then bank, then click Tampilkan for each combination")
        city_index = 0  # Start with first city (index 0)
        is_first_bank_in_city = True
        done = False
        
        # Get the first city (starts iteration after checkbox)
        print(f"\n[INFO] Getting first city (index {city_index})...")
        current_city = self._get_city_by_index(city_index)
        if not current_city:
            print(f"  [WARNING] No cities found at index {city_index}. This might be a timing issue or no cities available.")
            print(f"  [INFO] Scraping completed!")
            done = True
        
        while not done:
            if not current_city:
                print(f"  [INFO] No more cities. Scraping completed!")
                done = True
                break
            
            # Get all banks for this city (only once per city)
            # Wait a bit after city selection to ensure dropdown is ready
            time.sleep(2.0)  # Wait for city selection to fully load
            print(f"\n{'='*60}")
            print(f"[CITY] Processing: {current_city}")
            print(f"{'='*60}")
            bank_names = self._get_all_bank_names(city_index, city_already_selected=True)
            print(f"  [INFO] Found {len(bank_names)} banks in {current_city}")
            
            # If no banks found, wait a bit and retry once
            if not bank_names:
                print(f"  [WARNING] No banks found on first attempt, waiting and retrying...")
                time.sleep(3.0)
                bank_names = self._get_all_bank_names(city_index, city_already_selected=True)
                print(f"  [INFO] Found {len(bank_names)} banks in {current_city} (retry)")
            
            if not bank_names:
                print(f"  [WARNING] No banks found in {current_city}, moving to next city...")
                city_index += 1
                current_city = self._get_city_by_index(city_index)
                is_first_bank_in_city = True
                continue
                
            # Iterate through all banks in this city
            for bank_index, current_bank in enumerate(bank_names):
                print(f"\n  [BANK ({bank_index+1}/{len(bank_names)})] Processing: {current_bank}")
                
                # Select the bank by index (bank_index = 0 selects first span, bank_index = 1 selects second, etc.)
                # Add retry logic for index 0, especially after city change
                max_retries = 3 if bank_index == 0 else 1
                selected_bank_name = None
                
                for retry in range(max_retries):
                    selected_bank_name = self._select_bank_by_index(bank_index, city_index, city_already_selected=True)
                    if selected_bank_name:
                        break
                    elif retry < max_retries - 1:
                        print(f"  [WARNING] Could not select bank at index {bank_index}, retrying ({retry+1}/{max_retries})...")
                        time.sleep(2.0)  # Wait longer before retry
                
                if not selected_bank_name:
                    print(f"  [WARNING] Could not select bank at index {bank_index} after {max_retries} attempts")
                    continue
                
                time.sleep(2.0)  # Wait longer for bank selection to load (slower speed)
                
                # Click Tampilkan and extract data - MUST complete before moving to next bank
                print(f"  [INFO] Clicking Tampilkan and waiting for data extraction to complete...")
                extracted_data = self._click_tampilkan_and_extract_data(year, current_city, selected_bank_name)
                
                if extracted_data:
                    # Append to Excel - this must complete before moving to next bank
                    print(f"  [INFO] Appending data to Excel...")
                    self._append_to_excel(extracted_data, year, current_city, selected_bank_name, is_first_bank_in_city)
                    print(f"  [OK] Data successfully extracted and saved to Excel for {current_city} - {selected_bank_name}")
                    is_first_bank_in_city = False
                    # Wait longer to ensure Excel data is fully written
                    time.sleep(1.0)
                else:
                    print(f"  [WARNING] Failed to extract data for {current_city} - {selected_bank_name}")
                
                # Check if this is the last bank - if so, wait a bit before moving to next city
                if bank_index == len(bank_names) - 1:
                    print(f"  [INFO] This is the last bank ({bank_index+1}/{len(bank_names)}) in {current_city}")
                    time.sleep(1.0)  # Additional wait after last bank
            
            # After processing ALL banks in this city, move to next city
            print(f"  [INFO] Finished processing all {len(bank_names)} banks in {current_city}, moving to next city...")
            time.sleep(1.0)  # Wait before changing city
            city_index += 1
            current_city = self._get_city_by_index(city_index)
            is_first_bank_in_city = True
        
        # Finalize Excel file
        self._finalize_excel(year)
        
        print("\n[OK] All data collection completed!")
        print("[INFO] Closing browser...")
        self.cleanup()
    
    def _tick_checkboxes(self):
        """
        Find treeview elements and check the first 3 checkboxes
        This can be done in parallel with other operations for faster execution
        """
        print("  [INFO] Finding treeview elements and checking checkboxes...")
        time.sleep(0.5)  # Wait a bit for treeview to be ready
        
        treeview_ids = [
            "treeview-1012-record-BPK-901-000001"
        ]
        
        for treeview_id in treeview_ids:
            try:
                print(f"    [INFO] Looking for treeview element: {treeview_id}")
                
                # Wait for treeview element to be present
                wait = WebDriverWait(self.driver, 10)
                treeview_element = wait.until(
                    EC.presence_of_element_located((By.ID, treeview_id))
                )
                print(f"    [OK] Found treeview element: {treeview_id}")
                
                # Find nested divs with role="checkbox" inside this treeview element
                # Try multiple XPath patterns to find checkboxes
                checkboxes = []
                
                # Pattern 1: div with role="checkbox"
                checkboxes = treeview_element.find_elements(By.XPATH, ".//div[@role='checkbox']")
                if checkboxes:
                    print(f"    [DEBUG] Found {len(checkboxes)} checkbox(es) using pattern: div[@role='checkbox']")
                else:
                    # Pattern 2: input type="checkbox"
                    checkboxes = treeview_element.find_elements(By.XPATH, ".//input[@type='checkbox']")
                    if checkboxes:
                        print(f"    [DEBUG] Found {len(checkboxes)} checkbox(es) using pattern: input[@type='checkbox']")
                    else:
                        # Pattern 3: elements with checkbox in class
                        checkboxes = treeview_element.find_elements(By.XPATH, ".//div[contains(@class, 'checkbox')]")
                        if checkboxes:
                            print(f"    [DEBUG] Found {len(checkboxes)} checkbox(es) using pattern: div[contains(@class, 'checkbox')]")
                        else:
                            # Pattern 4: any element with aria-checked attribute
                            checkboxes = treeview_element.find_elements(By.XPATH, ".//*[@aria-checked]")
                            if checkboxes:
                                print(f"    [DEBUG] Found {len(checkboxes)} checkbox(es) using pattern: *[@aria-checked]")
                            else:
                                # Pattern 5: look for elements with checkbox-related attributes
                                checkboxes = treeview_element.find_elements(By.XPATH, ".//*[contains(@class, 'x-tree-checkbox') or contains(@class, 'tree-checkbox')]")
                                if checkboxes:
                                    print(f"    [DEBUG] Found {len(checkboxes)} checkbox(es) using pattern: *[contains(@class, 'tree-checkbox')]")
                                else:
                                    # Pattern 6: look for any clickable element that might be a checkbox
                                    checkboxes = treeview_element.find_elements(By.XPATH, ".//*[@role='checkbox' or @type='checkbox' or contains(@class, 'checkbox')]")
                                    if checkboxes:
                                        print(f"    [DEBUG] Found {len(checkboxes)} checkbox(es) using pattern: *[@role='checkbox' or @type='checkbox' or contains(@class, 'checkbox')]")
                
                if checkboxes:
                    print(f"    [OK] Found {len(checkboxes)} checkbox(es) in {treeview_id}")
                    for i, checkbox in enumerate(checkboxes):
                        try:
                            # Check if element is visible
                            if not checkbox.is_displayed():
                                print(f"    [DEBUG] Checkbox {i+1} is not visible, skipping")
                                continue
                            
                            # Get checkbox attributes for debugging
                            role = checkbox.get_attribute("role")
                            aria_checked = checkbox.get_attribute("aria-checked")
                            checkbox_type = checkbox.get_attribute("type")
                            
                            # Check if checkbox is already checked
                            if aria_checked == "true" or (checkbox_type == "checkbox" and checkbox.is_selected()):
                                print(f"    [INFO] Checkbox {i+1} already checked in {treeview_id}")
                                continue
                            
                            # Scroll to checkbox and click
                            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", checkbox)
                            time.sleep(0.5)  # Wait for scroll
                            
                            # Try clicking with JavaScript first
                            try:
                                self.driver.execute_script("arguments[0].click();", checkbox)
                                print(f"    [OK] Checked checkbox {i+1} in {treeview_id} (JavaScript click)")
                            except:
                                # Fallback to regular click
                                checkbox.click()
                                print(f"    [OK] Checked checkbox {i+1} in {treeview_id} (regular click)")
                            
                            time.sleep(0.5)  # Wait after click
                        except Exception as e:
                            print(f"    [WARNING] Could not check checkbox {i+1} in {treeview_id}: {e}")
                else:
                    print(f"    [WARNING] No checkboxes found in {treeview_id}")
            except Exception as e:
                print(f"    [WARNING] Could not find treeview element {treeview_id}: {e}")
        
        print("  [OK] Completed checkbox selection")
    
    def _select_initial_dropdowns_and_checkboxes(self):
        """
        3-step sequential process:
        1. Click dropdown arrow ext-gen1064 and select topmost <li>
        2. Click dropdown arrow ext-gen1069 and select topmost <li>
        3. Find treeview elements and check checkboxes inside them
        4. Click Tampilkan button and extract data
        
        Args:
            year: Selected year for data extraction
        """
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        
        # Step 1: Skip dropdown selections (handled in main loop)
        print("\n  [Step 4.1] Skipping city dropdown selection (handled in main loop)...")
        time.sleep(0.2)
        
        # Step 2: Skip bank dropdown selection (handled in main loop)
        print("\n  [Step 4.2] Skipping bank dropdown selection (handled in main loop)...")
        time.sleep(0.2)
        
        # Step 3: Find treeview element and check only the first checkbox
        # All three data types (Kredit, Total Aset, DPK) use the same checkbox
        print("\n  [Step 4.3] Finding treeview element and checking checkbox...")
        time.sleep(0.5)  # Wait a bit longer for treeview to be ready
        
        treeview_id = "treeview-1012-record-BPK-901-000001"
        
        try:
            print(f"    [INFO] Looking for treeview element: {treeview_id}")
            
            # Wait for treeview element to be present
            wait = WebDriverWait(self.driver, 10)
            treeview_element = wait.until(
                EC.presence_of_element_located((By.ID, treeview_id))
            )
            print(f"    [OK] Found treeview element: {treeview_id}")
            
            # Find nested divs with role="checkbox" inside this treeview element
            # Try multiple XPath patterns to find checkboxes
            checkboxes = []
            
            # Pattern 1: div with role="checkbox"
            checkboxes = treeview_element.find_elements(By.XPATH, ".//div[@role='checkbox']")
            if checkboxes:
                print(f"    [DEBUG] Found {len(checkboxes)} checkbox(es) using pattern: div[@role='checkbox']")
            else:
                # Pattern 2: input type="checkbox"
                checkboxes = treeview_element.find_elements(By.XPATH, ".//input[@type='checkbox']")
                if checkboxes:
                    print(f"    [DEBUG] Found {len(checkboxes)} checkbox(es) using pattern: input[@type='checkbox']")
                else:
                    # Pattern 3: elements with checkbox in class
                    checkboxes = treeview_element.find_elements(By.XPATH, ".//div[contains(@class, 'checkbox')]")
                    if checkboxes:
                        print(f"    [DEBUG] Found {len(checkboxes)} checkbox(es) using pattern: div[contains(@class, 'checkbox')]")
                    else:
                        # Pattern 4: any element with aria-checked attribute
                        checkboxes = treeview_element.find_elements(By.XPATH, ".//*[@aria-checked]")
                        if checkboxes:
                            print(f"    [DEBUG] Found {len(checkboxes)} checkbox(es) using pattern: *[@aria-checked]")
                        else:
                            # Pattern 5: look for elements with checkbox-related attributes
                            checkboxes = treeview_element.find_elements(By.XPATH, ".//*[contains(@class, 'x-tree-checkbox') or contains(@class, 'tree-checkbox')]")
                            if checkboxes:
                                print(f"    [DEBUG] Found {len(checkboxes)} checkbox(es) using pattern: *[contains(@class, 'tree-checkbox')]")
                            else:
                                # Pattern 6: look for any clickable element that might be a checkbox
                                checkboxes = treeview_element.find_elements(By.XPATH, ".//*[@role='checkbox' or @type='checkbox' or contains(@class, 'checkbox')]")
                                if checkboxes:
                                    print(f"    [DEBUG] Found {len(checkboxes)} checkbox(es) using pattern: *[@role='checkbox' or @type='checkbox' or contains(@class, 'checkbox')]")
            
            if checkboxes:
                print(f"    [OK] Found {len(checkboxes)} checkbox(es) in {treeview_id}")
                # Only check the first checkbox (all three data types use the same checkbox)
                checkbox = checkboxes[0]
                try:
                    # Check if element is visible
                    if not checkbox.is_displayed():
                        print(f"    [DEBUG] Checkbox is not visible, skipping")
                    else:
                        # Get checkbox attributes for debugging
                        aria_checked = checkbox.get_attribute("aria-checked")
                        checkbox_type = checkbox.get_attribute("type")
                        
                        # Check if checkbox is already checked
                        if aria_checked == "true" or (checkbox_type == "checkbox" and checkbox.is_selected()):
                            print(f"    [INFO] Checkbox already checked in {treeview_id}")
                        else:
                            # Scroll to checkbox and click
                            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", checkbox)
                            time.sleep(0.5)  # Wait for scroll
                            
                            # Try clicking with JavaScript first
                            try:
                                self.driver.execute_script("arguments[0].click();", checkbox)
                                print(f"    [OK] Checked checkbox in {treeview_id} (JavaScript click)")
                            except:
                                # Fallback to regular click
                                checkbox.click()
                                print(f"    [OK] Checked checkbox in {treeview_id} (regular click)")
                            
                            time.sleep(0.5)  # Wait after click
                except Exception as e:
                    print(f"    [WARNING] Could not check checkbox in {treeview_id}: {e}")
            else:
                print(f"    [WARNING] No checkboxes found in {treeview_id}")
        except Exception as e:
            print(f"    [WARNING] Could not find treeview element {treeview_id}: {e}")
        
        print("  [OK] Completed initial setup (dropdowns and checkbox)")
    
    def _get_city_by_index(self, index: int) -> str:
        """Get city name by index from dropdown ext-gen1064. Returns city name or None if not found."""
        try:
            print(f"    [DEBUG] Attempting to get city at index {index}...")
            # Click dropdown trigger to open
            dropdown_trigger = self.driver.find_element(By.ID, "ext-gen1064")
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown_trigger)
            time.sleep(0.5)
            self.driver.execute_script("arguments[0].click();", dropdown_trigger)
            print(f"    [DEBUG] Clicked city dropdown trigger")
            
            # Wait for dropdown to appear
            time.sleep(0.5)
            wait = WebDriverWait(self.driver, 10)
            wait.until(EC.presence_of_element_located((By.XPATH, "//ul[contains(@class, 'x-list-plain')]")))
            print(f"    [DEBUG] City dropdown appeared")
            
            # Get all city options
            li_elements = self.driver.find_elements(By.XPATH, "//li[@role='option' or contains(@class, 'x-boundlist-item')]")
            if not li_elements:
                li_elements = self.driver.find_elements(By.XPATH, "//ul[contains(@class, 'x-list-plain')]//li")
            
            print(f"    [DEBUG] Found {len(li_elements)} <li> elements in city dropdown")
            
            # Filter out empty elements and get by index
            valid_cities = [li for li in li_elements if li.text.strip()]
            print(f"    [DEBUG] Found {len(valid_cities)} valid cities (non-empty)")
            
            if valid_cities:
                for i, city_li in enumerate(valid_cities[:5]):  # Print first 5 for debugging
                    print(f"    [DEBUG] City {i}: '{city_li.text.strip()[:50]}...'")
            
            if index < len(valid_cities):
                city_name = valid_cities[index].text.strip()
                print(f"    [DEBUG] Selecting city at index {index}: '{city_name}'")
                # Select it
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", valid_cities[index])
                time.sleep(0.5)
                self.driver.execute_script("arguments[0].click();", valid_cities[index])
                time.sleep(2.0)  # Wait longer for PostBack and dropdown to update
                print(f"    [OK] Selected city: '{city_name}'")
                return city_name
            else:
                print(f"    [WARNING] Index {index} is out of range. Total valid cities: {len(valid_cities)}")
                # Close dropdown
                try:
                    from selenium.webdriver.common.keys import Keys
                    dropdown_trigger.send_keys(Keys.ESCAPE)
                except:
                    pass
                return None
        except Exception as e:
            print(f"    [ERROR] Could not get city by index {index}: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def _get_all_bank_names(self, city_index: int, city_already_selected: bool = False) -> list:
        """Get all bank names from dropdown ext-gen1069. Returns list of bank names.
        
        Args:
            city_index: Index of the city (used to ensure correct city is selected)
            city_already_selected: If True, skip city selection (city was already selected)
        """
        try:
            # Make sure we're on the correct city first (unless already selected)
            if not city_already_selected:
                current_city = self._get_city_by_index(city_index)
                if not current_city:
                    return []
                time.sleep(3.0)  # Wait longer for banks to load after city selection
            else:
                time.sleep(1.0)  # Wait a bit longer even if city already selected
            
            # Click dropdown trigger to open
            dropdown_trigger = self.driver.find_element(By.ID, "ext-gen1069")
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown_trigger)
            time.sleep(0.5)  # Longer wait before clicking
            self.driver.execute_script("arguments[0].click();", dropdown_trigger)
            
            # Wait for dropdown to appear
            time.sleep(0.5)
            wait = WebDriverWait(self.driver, 10)  # Increased wait time
            wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'x-boundlist')] | //table")))
            
            # Wait for tbody to be present and populated
            try:
                wait.until(EC.presence_of_element_located((By.XPATH, "//tbody[@id='treeview-1022-body']")))
                time.sleep(1.0)  # Wait for spans to be rendered
            except:
                pass
            
            # Get all spans with class="x-tree-node-text" within tbody id="treeview-1022-body"
            # This is the bank dropdown, not the checkbox area
            # THE FIX: Only use visible spans for indexing - filter immediately
            # Ensure we ONLY check spans inside treeview-1022-body tbody
            all_spans = self.driver.find_elements(By.XPATH, "//tbody[@id='treeview-1022-body']//span[@class='x-tree-node-text' or contains(@class, 'x-tree-node-text')]")
            
            # Filter to only visible, non-empty spans immediately
            # The XPath already ensures we're only getting spans from treeview-1022-body
            span_elements = [
                s for s in all_spans
                if s.is_displayed() and s.text.strip()
            ]
            
            bank_names = []
            skip_labels = ["Laporan Posisi Keuangan", "Laporan Laba Rugi", "Laporan", "Posisi", "Keuangan", "Laba", "Rugi"]
            
            if span_elements:
                for span in span_elements:
                    span_text = span.text.strip()
                    
                    # Skip if it's a known label (not a bank name)
                    is_label = any(label.lower() in span_text.lower() for label in skip_labels)
                    if is_label:
                        continue
                    
                    # Bank names typically contain numbers (bank codes)
                    has_number = any(char.isdigit() for char in span_text)
                    
                    # Only add if it looks like a bank name (has number or is reasonably long)
                    if has_number or len(span_text) > 15:
                        bank_names.append(span_text)
                        print(f"    [DEBUG] Found bank: '{span_text[:50]}...'")
            
            # Close dropdown
            try:
                from selenium.webdriver.common.keys import Keys
                dropdown_trigger.send_keys(Keys.ESCAPE)
            except:
                pass
            
            return bank_names
        except Exception as e:
            print(f"    [DEBUG] Could not get all bank names: {e}")
            return []
    
    def _select_bank_by_index(self, bank_index: int, city_index: int, city_already_selected: bool = False) -> str:
        """Select a bank by its index from dropdown ext-gen1069. 
        
        Args:
            bank_index: Index of the bank to select (0 = first span, 1 = second span, etc.)
            city_index: Index of the city (used to ensure correct city is selected)
            city_already_selected: If True, skip city selection (city was already selected)
        
        Returns:
            Bank name if successful, empty string if failed.
        """
        try:
            # Make sure we're on the correct city first (unless already selected)
            if not city_already_selected:
                current_city = self._get_city_by_index(city_index)
                if not current_city:
                    return ""
                time.sleep(3.0)  # Wait longer for banks to load after city selection
            else:
                # If this is index 0 (first bank), wait longer to ensure dropdown is ready
                if bank_index == 0:
                    time.sleep(2.0)  # Longer wait for first bank after city change
                else:
                    time.sleep(1.0)  # Wait between bank selections
            
            # Click dropdown trigger to open
            dropdown_trigger = self.driver.find_element(By.ID, "ext-gen1069")
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown_trigger)
            time.sleep(0.3)
            self.driver.execute_script("arguments[0].click();", dropdown_trigger)
            
            # Wait for dropdown to appear
            time.sleep(0.5)
            wait = WebDriverWait(self.driver, 10)
            wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'x-boundlist')] | //table")))
            
            # Wait for tbody to be present and populated
            # For index 0 (first bank), wait longer to ensure dropdown is fully loaded
            wait_time = 1.5 if bank_index == 0 else 0.5
            try:
                wait.until(EC.presence_of_element_located((By.XPATH, "//tbody[@id='treeview-1022-body']")))
                time.sleep(wait_time)  # Longer wait for first bank, shorter for others
            except:
                pass
            
            # Additional wait for spans to be rendered, especially for index 0
            if bank_index == 0:
                time.sleep(1.0)
            
            # Get all spans with class="x-tree-node-text" within tbody id="treeview-1022-body"
            # This ensures we're clicking the dropdown tr, not the checkbox tr
            # THE FIX: Only use visible spans for indexing - filter immediately
            # Ensure we ONLY check spans inside treeview-1022-body tbody
            all_spans = self.driver.find_elements(By.XPATH, "//tbody[@id='treeview-1022-body']//span[@class='x-tree-node-text' or contains(@class, 'x-tree-node-text')]")
            
            # Filter to only visible, non-empty spans immediately
            # The XPath already ensures we're only getting spans from treeview-1022-body
            span_elements = [
                s for s in all_spans
                if s.is_displayed() and s.text.strip()
            ]
            
            print(f"    [DEBUG] Found {len(span_elements)} visible, non-empty span elements in dropdown (from treeview-1022-body)")
            
            # Filter spans the same way as _get_all_bank_names
            skip_labels = ["Laporan Posisi Keuangan", "Laporan Laba Rugi", "Laporan", "Posisi", "Keuangan", "Laba", "Rugi"]
            valid_bank_spans = []
            
            for span in span_elements:
                try:
                    span_text = span.text.strip()
                    
                    # Skip if it's a known label (not a bank name)
                    is_label = any(label.lower() in span_text.lower() for label in skip_labels)
                    if is_label:
                        continue
                    
                    # Bank names typically contain numbers (bank codes)
                    has_number = any(char.isdigit() for char in span_text)
                    
                    # Only add if it looks like a bank name (has number or is reasonably long)
                    if has_number or len(span_text) > 15:
                        valid_bank_spans.append((span, span_text))
                        print(f"    [DEBUG] Valid bank span {len(valid_bank_spans)-1}: '{span_text[:50]}...'")
                except:
                    # Element might be stale, skip it
                    continue
            
            print(f"    [DEBUG] Total valid bank spans: {len(valid_bank_spans)}")
            
            # Select by index
            if bank_index < len(valid_bank_spans):
                selected_span, bank_name = valid_bank_spans[bank_index]
                
                # Find the parent tr within the tbody (this is the dropdown row, not checkbox)
                try:
                    # Find the closest tr ancestor (should be within treeview-1022-body)
                    parent_tr = selected_span.find_element(By.XPATH, "./ancestor::tr[1]")
                    clickable_elem = parent_tr
                except:
                    # Fallback: try to find any clickable ancestor
                    try:
                        clickable_elem = selected_span.find_element(By.XPATH, "./ancestor::*[@role='row' or contains(@class, 'x-boundlist-item')][1]")
                    except:
                        clickable_elem = selected_span
                
                # Select it
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", clickable_elem)
                time.sleep(0.3)
                self.driver.execute_script("arguments[0].click();", clickable_elem)
                time.sleep(0.5)  # Wait for PostBack
                print(f"    [DEBUG] Selected bank at index {bank_index}: '{bank_name[:50]}...'")
                return bank_name
            else:
                # Close dropdown if index out of range
                try:
                    from selenium.webdriver.common.keys import Keys
                    dropdown_trigger.send_keys(Keys.ESCAPE)
                except:
                    pass
                print(f"    [WARNING] Bank index {bank_index} is out of range. Total valid banks: {len(valid_bank_spans)}")
                return ""
        except Exception as e:
            print(f"    [DEBUG] Could not select bank by index {bank_index}: {e}")
            return ""
    
    def _get_bank_by_index(self, city_index: int, bank_index: int, city_already_selected: bool = False) -> str:
        """Get bank name by index from dropdown ext-gen1069. Returns bank name or None if not found.
        
        Args:
            city_index: Index of the city (used to ensure correct city is selected)
            bank_index: Index of the bank to select
            city_already_selected: If True, skip city selection (city was already selected)
        """
        # Get all bank names first
        bank_names = self._get_all_bank_names(city_index, city_already_selected)
        
        if bank_index < len(bank_names):
            bank_name = bank_names[bank_index]
            # Select the bank by name
            if self._select_bank_by_name(bank_name):
                return bank_name
        
        return None
    
    def _get_city_and_bank_by_index(self, city_index: int, bank_index: int) -> tuple:
        """Get both city and bank by their indices. Returns (city_name, bank_name) or (None, None)"""
        city = self._get_city_by_index(city_index)
        if not city:
            return (None, None)
        
        # City is already selected, so pass city_already_selected=True
        bank = self._get_bank_by_index(city_index, bank_index, city_already_selected=True)
        if not bank:
            return (city, None)
        
        return (city, bank)
    
    def _select_next_city(self) -> bool:
        """Select next city from dropdown ext-gen1064. Returns True if successful, False if last city."""
        try:
            # Click dropdown trigger
            dropdown_trigger = self.driver.find_element(By.ID, "ext-gen1064")
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown_trigger)
            time.sleep(0.5)
            self.driver.execute_script("arguments[0].click();", dropdown_trigger)
            
            # Wait for dropdown to appear
            time.sleep(0.5)
            wait = WebDriverWait(self.driver, 5)
            wait.until(EC.presence_of_element_located((By.XPATH, "//ul[contains(@class, 'x-list-plain')]")))
            
            # Get current selected city
            current_city = self._get_current_selected_city()
            
            # Get all city options
            li_elements = self.driver.find_elements(By.XPATH, "//li[@role='option' or contains(@class, 'x-boundlist-item')]")
            if not li_elements:
                li_elements = self.driver.find_elements(By.XPATH, "//ul[contains(@class, 'x-list-plain')]//li")
            
            # Find current city index and select next
            current_index = -1
            for i, li in enumerate(li_elements):
                li_text = li.text.strip()
                if li_text and (current_city.lower() in li_text.lower() or li_text.lower() in current_city.lower()):
                    current_index = i
                    break
            
            # Select next city
            if current_index >= 0 and current_index < len(li_elements) - 1:
                next_li = li_elements[current_index + 1]
                next_text = next_li.text.strip()
                if next_text:
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_li)
                    time.sleep(0.5)
                    self.driver.execute_script("arguments[0].click();", next_li)
                    print(f"  [OK] Selected next city: '{next_text}'")
                    time.sleep(0.5)
                    return True
            else:
                print(f"  [INFO] Already at last city")
                # Close dropdown by clicking outside or pressing ESC
                try:
                    from selenium.webdriver.common.keys import Keys
                    dropdown_trigger.send_keys(Keys.ESCAPE)
                except:
                    pass
                return False
        except Exception as e:
            print(f"  [WARNING] Could not select next city: {e}")
            return False
    
    def _select_next_bank(self) -> bool:
        """Select next bank from dropdown ext-gen1069. Returns True if successful, False if last bank."""
        try:
            # Click dropdown trigger
            dropdown_trigger = self.driver.find_element(By.ID, "ext-gen1069")
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown_trigger)
            time.sleep(0.5)
            self.driver.execute_script("arguments[0].click();", dropdown_trigger)
            
            # Wait for dropdown to appear
            time.sleep(0.5)
            wait = WebDriverWait(self.driver, 5)
            wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'x-boundlist')] | //table")))
            
            # Get current selected bank
            current_bank = self._get_current_selected_bank()
            
            # Get all bank options (try <tr> first, then <li>)
            tr_elements = self.driver.find_elements(By.XPATH, "//div[contains(@class, 'x-boundlist')]//tr | //div[contains(@class, 'x-layer')]//tr")
            if not tr_elements:
                tr_elements = self.driver.find_elements(By.XPATH, "//tr")
            
            elements = tr_elements if tr_elements else []
            if not elements:
                # Fallback to <li>
                li_elements = self.driver.find_elements(By.XPATH, "//li[@role='option' or contains(@class, 'x-boundlist-item')]")
                if not li_elements:
                    li_elements = self.driver.find_elements(By.XPATH, "//ul[contains(@class, 'x-list-plain')]//li")
                elements = li_elements
            
            # Find current bank index and select next
            current_index = -1
            for i, elem in enumerate(elements):
                if not elem.is_displayed():
                    continue
                elem_text = elem.text.strip()
                if elem_text and (current_bank.lower() in elem_text.lower() or elem_text.lower() in current_bank.lower()):
                    current_index = i
                    break
            
            # Select next bank
            if current_index >= 0 and current_index < len(elements) - 1:
                next_elem = elements[current_index + 1]
                next_text = next_elem.text.strip()
                if next_text and next_elem.is_displayed():
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_elem)
                    time.sleep(0.5)
                    self.driver.execute_script("arguments[0].click();", next_elem)
                    print(f"  [OK] Selected next bank: '{next_text[:50]}...'")
                    time.sleep(0.5)
                    return True
            else:
                print(f"  [INFO] Already at last bank")
                # Close dropdown
                try:
                    from selenium.webdriver.common.keys import Keys
                    dropdown_trigger.send_keys(Keys.ESCAPE)
                except:
                    pass
                return False
        except Exception as e:
            print(f"  [WARNING] Could not select next bank: {e}")
            return False
    
    def _select_first_bank(self) -> bool:
        """Select first bank from dropdown ext-gen1069. Returns True if successful, False if no banks."""
        try:
            # Click dropdown trigger
            dropdown_trigger = self.driver.find_element(By.ID, "ext-gen1069")
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown_trigger)
            time.sleep(0.5)
            self.driver.execute_script("arguments[0].click();", dropdown_trigger)
            
            # Wait for dropdown to appear
            time.sleep(0.5)
            wait = WebDriverWait(self.driver, 5)
            wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'x-boundlist')] | //table")))
            
            # Get all bank options using <span class="x-tree-node-text">
            # But only within tr elements (to avoid selecting labels/headers)
            # For first bank, use index 0 (first valid bank span)
            span_elements = self.driver.find_elements(By.XPATH, "//tr//span[@class='x-tree-node-text']")
            
            if not span_elements:
                # Fallback: try with contains
                span_elements = self.driver.find_elements(By.XPATH, "//tr//span[contains(@class, 'x-tree-node-text')]")
            
            if span_elements:
                # Filter out empty and non-visible spans, and filter out labels
                valid_banks = []
                skip_labels = ["Laporan Posisi Keuangan", "Laporan Laba Rugi", "Laporan", "Posisi", "Keuangan", "Laba", "Rugi"]
                
                for i, span in enumerate(span_elements):
                    if not span.is_displayed():
                        continue
                    span_text = span.text.strip()
                    if not span_text:
                        continue
                    
                    # Skip if it's a known label (not a bank name)
                    is_label = any(label.lower() in span_text.lower() for label in skip_labels)
                    if is_label:
                        continue
                    
                    # Bank names typically contain numbers
                    has_number = any(char.isdigit() for char in span_text)
                    
                    if has_number or len(span_text) > 20:  # Bank names are usually longer
                        valid_banks.append((i, span, span_text))
                
                # Select first bank (index 0)
                if len(valid_banks) > 0:
                    _, bank_span, bank_name = valid_banks[0]
                    # Find the parent tr or clickable element to click
                    try:
                        parent_tr = bank_span.find_element(By.XPATH, "./ancestor::tr[1]")
                        clickable_elem = parent_tr
                    except:
                        try:
                            clickable_elem = bank_span.find_element(By.XPATH, "./ancestor::*[@role='row' or contains(@class, 'x-boundlist-item')][1]")
                        except:
                            clickable_elem = bank_span
                    
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", clickable_elem)
                    time.sleep(0.5)
                    self.driver.execute_script("arguments[0].click();", clickable_elem)
                    print(f"  [OK] Selected first bank (index 0): '{bank_name[:50]}...'")
                    time.sleep(0.5)
                    return True
            else:
                # Fallback: try tr elements
                tr_elements = self.driver.find_elements(By.XPATH, "//tr[@data-boundview='treeview-1022']")
                if tr_elements:
                    for i, tr in enumerate(tr_elements):
                        if not tr.is_displayed():
                            continue
                        tr_text = tr.text.strip()
                        if tr_text:
                            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", tr)
                            time.sleep(0.5)
                            self.driver.execute_script("arguments[0].click();", tr)
                            print(f"  [OK] Selected first bank (fallback, index {i}): '{tr_text[:50]}...'")
                            time.sleep(0.5)
                            return True
            
            return False
        except Exception as e:
            print(f"  [WARNING] Could not select first bank: {e}")
            return False
    
    def _click_tampilkan_and_extract_data(self, year: str, city: str, bank: str) -> dict:
        """Click Tampilkan button, wait for report, and extract data"""
        try:
            # Make sure we're on default content and close any open dropdowns
            self.driver.switch_to.default_content()
            time.sleep(0.3)
            
            # Close any open dropdowns by pressing ESC
            try:
                from selenium.webdriver.common.keys import Keys
                body = self.driver.find_element(By.TAG_NAME, "body")
                body.send_keys(Keys.ESCAPE)
                time.sleep(0.3)
            except:
                pass
            
            # Click Tampilkan button
            tampilkan_button = None
            try:
                tampilkan_button = self.driver.find_element(By.ID, "ShowReportButton-btnInnerEl")
            except:
                tampilkan_button = self.driver.find_element(By.XPATH, "//span[contains(text(), 'Tampilkan')]")
            
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", tampilkan_button)
            time.sleep(0.5)
            self.driver.execute_script("arguments[0].click();", tampilkan_button)
            print(f"  [OK] Clicked 'Tampilkan' button")
            
            # Wait for report to load
            print(f"  [INFO] Waiting for report to load...")
            self._wait_for_report_loaded(max_wait=30)
            
            # Extract data - use the provided city and bank names (from dropdown)
            extracted_data = self._extract_report_data(year, city, bank)
            return extracted_data
        except Exception as e:
            print(f"  [ERROR] Error in _click_tampilkan_and_extract_data: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def _initialize_excel(self, year: str):
        """Initialize Excel workbook and worksheet"""
        if not Workbook:
            print("  [ERROR] openpyxl not installed. Cannot create Excel file.")
            return
        
        self.excel_wb = Workbook()
        self.excel_ws = self.excel_wb.active
        self.excel_ws.title = f"OJK Report {year}"
        self.excel_row = 1
        
        # Set column widths
        self.excel_ws.column_dimensions['A'].width = 30
        self.excel_ws.column_dimensions['B'].width = 30
        self.excel_ws.column_dimensions['C'].width = 20
        self.excel_ws.column_dimensions['D'].width = 20
        
        print(f"  [OK] Excel file initialized")
    
    def _append_to_excel(self, data: dict, year: str, city: str, bank: str, is_first_bank_in_city: bool):
        """Append extracted data to Excel worksheet"""
        if not self.excel_wb or not self.excel_ws:
            print("  [ERROR] Excel not initialized")
            return
        
        try:
            from openpyxl.styles import Font
            
            previous_year = str(int(year) - 1)
            
            # Add blank row if not first bank in city
            if not is_first_bank_in_city:
                self.excel_row += 1
            
            # Row: Kota/Kabupaten
            self.excel_ws[f'A{self.excel_row}'] = f"Kota/Kabupaten: {city}"
            self.excel_ws[f'A{self.excel_row}'].font = Font(bold=True)
            self.excel_row += 1
            
            # Row: Bank
            self.excel_ws[f'A{self.excel_row}'] = f"Bank: {bank}"
            self.excel_ws[f'A{self.excel_row}'].font = Font(bold=True)
            self.excel_row += 1
            
            # Row: Kredit
            kredit_current = data.get(f'Kredit {year}', 0)
            kredit_previous = data.get(f'Kredit {previous_year}', 0)
            self.excel_ws[f'A{self.excel_row}'] = f"Kredit {year}:"
            self.excel_ws[f'B{self.excel_row}'] = f"{kredit_current:,.0f}".replace(',', '.') if isinstance(kredit_current, (int, float)) else str(kredit_current)
            self.excel_ws[f'C{self.excel_row}'] = f"Kredit {previous_year}:"
            self.excel_ws[f'D{self.excel_row}'] = f"{kredit_previous:,.0f}".replace(',', '.') if isinstance(kredit_previous, (int, float)) else str(kredit_previous)
            self.excel_row += 1
            
            # Row: Total Aset
            total_aset_current = data.get(f'Total Aset {year}', 0)
            total_aset_previous = data.get(f'Total Aset {previous_year}', 0)
            self.excel_ws[f'A{self.excel_row}'] = f"Total Aset {year}:"
            self.excel_ws[f'B{self.excel_row}'] = f"{total_aset_current:,.0f}".replace(',', '.') if isinstance(total_aset_current, (int, float)) else str(total_aset_current)
            self.excel_ws[f'C{self.excel_row}'] = f"Total Aset {previous_year}:"
            self.excel_ws[f'D{self.excel_row}'] = f"{total_aset_previous:,.0f}".replace(',', '.') if isinstance(total_aset_previous, (int, float)) else str(total_aset_previous)
            self.excel_row += 1
            
            # Row: DPK
            dpk_current = data.get(f'DPK {year}', 0)
            dpk_previous = data.get(f'DPK {previous_year}', 0)
            self.excel_ws[f'A{self.excel_row}'] = f"DPK {year}:"
            self.excel_ws[f'B{self.excel_row}'] = f"{dpk_current:,.0f}".replace(',', '.') if isinstance(dpk_current, (int, float)) else str(dpk_current)
            self.excel_ws[f'C{self.excel_row}'] = f"DPK {previous_year}:"
            self.excel_ws[f'D{self.excel_row}'] = f"{dpk_previous:,.0f}".replace(',', '.') if isinstance(dpk_previous, (int, float)) else str(dpk_previous)
            self.excel_row += 1
            
            print(f"  [OK] Appended data to Excel (row {self.excel_row - 1})")
        except Exception as e:
            print(f"  [ERROR] Error appending to Excel: {e}")
            import traceback
            traceback.print_exc()
    
    def _finalize_excel(self, year: str):
        """Save and close Excel workbook"""
        if not self.excel_wb:
            print("  [ERROR] Excel not initialized")
            return
        
        try:
            filename = f"OJK_Report_{year}.xlsx"
            filepath = self.output_dir / filename
            self.excel_wb.save(filepath)
            print(f"  [OK] Excel file saved to {filepath}")
            print(f"  [OK] Total rows written: {self.excel_row - 1}")
        except Exception as e:
            print(f"  [ERROR] Error saving Excel file: {e}")
            import traceback
            traceback.print_exc()
    
    def _find_combo_name_by_keyword(self, keyword: str) -> str:
        """
        Find combobox name by keyword
        
        Args:
            keyword: Keyword to search for (e.g., "month", "year", "province")
            
        Returns:
            Component name if found, empty string otherwise
        """
        if self.extjs is None:
            return ""
        
        combos = self.extjs.list_all_combos()
        keyword_lower = keyword.lower()
        
        for combo in combos:
            name = (combo.get('name') or '').lower()
            id_val = (combo.get('id') or '').lower()
            input_id = (combo.get('inputId') or '').lower()
            
            if (keyword_lower in name or 
                keyword_lower in id_val or 
                keyword_lower in input_id):
                return combo.get('name', '')
        
        return ''
    
    def _wait_for_report_loaded(self, max_wait: int = 60) -> bool:
        """
        Wait for report to load by checking for required identifiers in the page
        
        Args:
            max_wait: Maximum time to wait in seconds (default 60)
            
        Returns:
            Always returns True after max_wait seconds (will create Excel with whatever data is found)
        """
        start_time = time.time()
        check_interval = 0.5  # Check every 0.5 second
        
        # Required identifiers to check for (for logging purposes)
        required_identifiers = [
            "Kredit",
            "DPK", 
            "Total Aset",
            "Laba Kotor",
            "Rasio"
        ]
        
        found_identifiers = set()
        
        while time.time() - start_time < max_wait:
            try:
                # Check main page source for report content
                self.driver.switch_to.default_content()
                page_source = self.driver.page_source
                page_source_lower = page_source.lower()
                
                # Check for all required identifiers (for logging)
                for identifier in required_identifiers:
                    if identifier not in found_identifiers:
                        if identifier in page_source or identifier.lower() in page_source_lower:
                            found_identifiers.add(identifier)
                            print(f"    [DEBUG] Found identifier: '{identifier}' ({len(found_identifiers)}/{len(required_identifiers)})")
                
                # Wait before next check
                elapsed = int(time.time() - start_time)
                remaining = max_wait - elapsed
                if remaining > 0 and remaining % 10 == 0:  # Print every 10 seconds
                    print(f"    [DEBUG] Waiting... ({remaining}s remaining, found {len(found_identifiers)}/{len(required_identifiers)} identifiers)")
                
                time.sleep(check_interval)
                
            except Exception as e:
                # If error occurs, wait and try again
                time.sleep(check_interval)
                continue
        
        # Always return True after waiting - will create Excel with whatever data is found
        print(f"    [DEBUG] Wait completed. Found identifiers: {found_identifiers}")
        return True
    
    def _extract_report_data(self, selected_year: str, city: str = None, bank: str = None) -> dict:
        """
        Extract financial data from the generated report
        
        Args:
            selected_year: The selected year (e.g., "2024")
            city: The selected city (optional, will try to extract from page if not provided)
            bank: The selected bank (optional, will try to extract from page if not provided)
            
        Returns:
            Dictionary with extracted data
        """
        try:
            # Try to find report in iframe first, then main page
            print("    [DEBUG] Checking for report in iframes...")
            iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
            page_source = None
            
            # Check iframes for report content
            report_iframe = None
            for iframe in iframes:
                try:
                    self.driver.switch_to.frame(iframe)
                    iframe_source = self.driver.page_source
                    if "Kredit" in iframe_source or "Aset" in iframe_source or "DPK" in iframe_source:
                        print("    [DEBUG] Found report content in iframe")
                        page_source = iframe_source
                        report_iframe = iframe  # Keep reference to the iframe
                        # Stay in iframe context for XPath searches
                        break
                    self.driver.switch_to.default_content()
                except:
                    self.driver.switch_to.default_content()
                    continue

            # If not in iframe, use main page
            if page_source is None:
                print("    [DEBUG] Using main page source")
                self.driver.switch_to.default_content()
                page_source = self.driver.page_source
            # else: Stay in iframe context - don't switch back yet
            
            # Identifiers to check for (at least one should be present for validation)
            identifiers_to_check = [
                "Kepada BPR",
                "Kepada Bank Umum",
                "pihak terkait",
                "pihak tidak terkait",
                "Total Aset",
                "Tabungan",
                "Deposito"
            ]
            
            def check_identifiers_in_soup(soup: BeautifulSoup, identifiers: list) -> tuple[bool, str]:
                """
                Check if identifiers exist in BeautifulSoup and have valid data records
                Simplified: just check if identifier exists and has next divs (data)
                
                Args:
                    soup: BeautifulSoup object to check
                    identifiers: List of identifier strings to check for
                    
                Returns:
                    Tuple of (found: bool, identifier_name: str)
                """
                for identifier in identifiers:
                    # Find the <div> whose text contains our identifier
                    # Be more specific: skip very long divs (entire page content)
                    for div in soup.find_all('div'):
                        text = div.get_text(strip=True)
                        if not text:
                            continue
                        
                        # Skip divs that are too long (likely contain entire page content)
                        if len(text) > 5000:
                            continue
                        
                        if identifier.lower() in text.lower():
                            # Prefer divs where identifier is a significant part of the text
                            text_lower = text.lower()
                            identifier_lower = identifier.lower()
                            
                            # If text is short or identifier is at the start/end, it's likely the right div
                            if len(text) < 200 or text_lower.startswith(identifier_lower) or text_lower.endswith(identifier_lower):
                                # Check if there are divs after this one (indicating data is present)
                                all_divs = soup.find_all('div')
                                try:
                                    div_index = all_divs.index(div)
                                    # If there are more divs after this identifier div, data is likely present
                                    if div_index < len(all_divs) - 1:
                                        return True, identifier
                                except ValueError:
                                    # div not in all_divs (shouldn't happen, but handle it)
                                    pass
                                break  # Found identifier, no need to continue searching
                return False, ""
            
            def refresh_page_source(report_iframe_ref) -> tuple[str, object]:
                """
                Refresh page source from iframe or main page
                
                Args:
                    report_iframe_ref: Reference to the iframe element (or None)
                    
                Returns:
                    Tuple of (page_source: str, updated_iframe_ref: object)
                """
                if report_iframe_ref:
                    # We're in iframe context, refresh iframe source
                    try:
                        self.driver.switch_to.frame(report_iframe_ref)
                        new_page_source = self.driver.page_source
                        return new_page_source, report_iframe_ref
                    except:
                        # Iframe might have changed, find it again
                        self.driver.switch_to.default_content()
                        iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
                        for iframe in iframes:
                            try:
                                self.driver.switch_to.frame(iframe)
                                iframe_source = self.driver.page_source
                                if "Kredit" in iframe_source or "Aset" in iframe_source or "DPK" in iframe_source:
                                    return iframe_source, iframe
                                self.driver.switch_to.default_content()
                            except:
                                self.driver.switch_to.default_content()
                                continue
                        # Fallback to main page if iframe not found
                        self.driver.switch_to.default_content()
                        return self.driver.page_source, None
                else:
                    # Using main page, refresh main page source
                    self.driver.switch_to.default_content()
                    return self.driver.page_source, None
            
            # Validation: Wait for page to fully load by checking if records are found by identifiers
            print("    [INFO] Memvalidasi halaman telah dimuat sepenuhnya (memeriksa identifier setiap 10 detik)...")
            max_wait_attempts = 20  # Maximum 20 attempts = 200 seconds (20 * 10s = 200s)
            wait_interval = 10  # Wait 10 seconds between checks
            page_fully_loaded = False
            soup = None
            
            for attempt in range(max_wait_attempts):
                # Always refresh page_source and re-parse BeautifulSoup to get latest content
                print(f"    [INFO] Percobaan {attempt + 1}/{max_wait_attempts}: Memperbarui page_source dan mem-parse ulang BeautifulSoup...")
                page_source, report_iframe = refresh_page_source(report_iframe)
                
                # Re-parse BeautifulSoup with fresh page_source
                soup = BeautifulSoup(page_source, 'html.parser')
                print(f"    [DEBUG] BeautifulSoup telah di-parse ulang (ukuran: {len(page_source)} karakter)")
                
                # Check if identifiers exist in the parsed BeautifulSoup
                record_found, found_identifier = check_identifiers_in_soup(soup, identifiers_to_check)
                
                if record_found:
                    page_fully_loaded = True
                    print(f"    [OK] Halaman telah dimuat sepenuhnya - Identifier '{found_identifier}' ditemukan dengan data (percobaan {attempt + 1})")
                    break
                else:
                    print(f"    [INFO] Halaman belum sepenuhnya dimuat - Identifier tidak ditemukan (percobaan {attempt + 1}/{max_wait_attempts})")
                    if attempt < max_wait_attempts - 1:  # Don't wait on last attempt
                        print(f"    [INFO] Menunggu {wait_interval} detik sebelum memeriksa lagi...")
                        time.sleep(wait_interval)
                    else:
                        print(f"    [WARNING] Mencapai batas maksimum percobaan ({max_wait_attempts})")
            
            if not page_fully_loaded:
                print("    [WARNING] Halaman mungkin belum sepenuhnya dimuat - Identifier tidak ditemukan setelah waktu tunggu maksimum")
                print("    [WARNING] Melanjutkan ekstraksi dengan data yang tersedia...")
                # Ensure soup is parsed with latest page_source
                if soup is None:
                    page_source, report_iframe = refresh_page_source(report_iframe)
                soup = BeautifulSoup(page_source, 'html.parser')
            
            # Helper function to split concatenated numbers (current year + previous year)
            def split_concatenated_numbers(text: str) -> tuple[str, str]:
                """
                Split concatenated numbers like "23,122,1223,112,122" into two numbers.
                Format: After comma, max 3 digits. If next is comma, same year. If next is digit, next year starts.
                
                Example: "23,122,1223,112,122" -> ("23,122,122", "3,112,122")
                
                Returns:
                    Tuple of (current_year_text, previous_year_text)
                """
                if not text or ',' not in text:
                    return text, ""
                
                import re
                
                # Find the split point: look for pattern where after comma, we have 3 digits, then a digit (not comma)
                # Pattern: comma, then exactly 3 digits, then a digit (not comma)
                # This indicates the start of the next year
                # Example: "23,122,1223" -> split after "122" (3 digits), before "3" (digit)
                pattern = r',(\d{3})(\d)(?=,|\d|$)'
                match = re.search(pattern, text)
                
                if match:
                    # Found split point: the digit after the 3-digit group starts the next year
                    # Split position is right before that digit
                    split_pos = match.end(1)  # Position after the 3 digits (before the next digit)
                    current_year_text = text[:split_pos].rstrip(',')
                    previous_year_text = text[split_pos:]
                    
                    # Remove leading comma from previous year if any
                    previous_year_text = previous_year_text.lstrip(',')
                    
                    print(f"      [DEBUG] Split '{text}' -> Current: '{current_year_text}', Previous: '{previous_year_text}'")
                    return current_year_text, previous_year_text
                
                # No split found, return original text as current year
                return text, ""
            
            # Helper function to extract numeric value from text
            def extract_number(text: str) -> float:
                """Extract an integer-like number from Indonesian-style formatted text."""
                if not text:
                    return 0.0
                
                original_text = text
                # Normalize spaces
                text = text.replace('\xa0', ' ').replace('&nbsp;', ' ').strip()
                
                # Keep digits only (we just need whole numbers)
                digits_only = re.sub(r'\D', '', text)
                if not digits_only:
                    print(f"      [DEBUG]     Failed to extract number from: '{original_text}'")
                    return 0.0
                
                value = float(digits_only)
                if original_text.strip() != digits_only:
                    print(f"      [DEBUG]     Raw: '{original_text}' -> Digits: '{digits_only}' -> Number: {value}")
                return value
            
            # Helper function to find identifier and get next 2 div values using BeautifulSoup
            def find_and_extract(identifier: str):
                """
                Find the div containing `identifier` and return the next 2 div values
                (current year, previous year) - simplified approach using inline divs
                """
                values = []
                
                # 1. Find the <div> whose text contains our identifier
                # Be more specific: look for divs where identifier is the main/only text, not in huge text blocks
                label_div = None
                all_divs = soup.find_all('div')
                
                for div in all_divs:
                    text = div.get_text(strip=True)
                    if not text:
                        continue
                    
                    # Skip divs that are too long (likely contain entire page content)
                    if len(text) > 5000:  # Skip very long divs (entire page content)
                        continue
                    
                    # Check if identifier is in text
                    if identifier.lower() in text.lower():
                        # Prefer divs where identifier is a significant part of the text
                        # or where text is relatively short (more specific match)
                        text_lower = text.lower()
                        identifier_lower = identifier.lower()
                        
                        # If text is short or identifier is at the start/end, it's likely the right div
                        if len(text) < 200 or text_lower.startswith(identifier_lower) or text_lower.endswith(identifier_lower):
                            label_div = div
                            print(f"    [DEBUG] Found identifier '{identifier}' in <div>: '{text[:100]}...' (length: {len(text)})")
                            break
                        # Otherwise, continue searching for a better match
                
                if not label_div:
                    print(f"    [DEBUG] Identifier '{identifier}' NOT FOUND in page")
                    return values
                
                # 2. Find the index of this div and get the next divs that contain numeric values
                try:
                    div_index = all_divs.index(label_div)
                    # Look through next divs to find numeric values (up to 10 divs ahead)
                    numeric_count = 0
                    for j in range(1, min(11, len(all_divs) - div_index)):  # Check up to 10 next divs
                        if numeric_count >= 2:  # We need 2 values
                            break
                            
                        if div_index + j < len(all_divs):
                            next_div = all_divs[div_index + j]
                            div_text = next_div.get_text(strip=True)
                            
                            if not div_text:
                                continue
                            
                            # Skip very long divs (likely contain entire page content)
                            if len(div_text) > 5000:
                                continue
                            
                            # Skip if it's clearly another identifier (contains common identifier keywords)
                            if any(keyword in div_text.lower() for keyword in ['kepada', 'pihak', 'bank', 'bpr', 'report viewer', 'configuration error']):
                                # This is likely another identifier or error message, skip it
                                continue
                            
                            # Check if this div contains numbers (might be formatted like "1.234.567" or "1,234,567")
                            # It might also contain concatenated numbers like "23,122,1223,112,122"
                            
                            # Check if it contains concatenated numbers (has comma and might have split pattern)
                            if ',' in div_text and len(div_text) > 10:
                                # Try to split concatenated numbers
                                current_year_text, prev_year_text = split_concatenated_numbers(div_text)
                                
                                if prev_year_text:
                                    # Found concatenated numbers, extract both
                                    current_number = extract_number(current_year_text)
                                    prev_number = extract_number(prev_year_text)
                                    
                                    # Validate numbers are reasonable
                                    if (current_number > 0 and current_number < 1e15 and current_number != float('inf') and
                                        prev_number >= 0 and prev_number < 1e15 and prev_number != float('inf')):
                                        
                                        if numeric_count == 0:
                                            # Add current year first
                                            print(f"    [DEBUG]   Next div[{j}] (Year {selected_year}): "
                                                  f"'{div_text}' -> Split: '{current_year_text}' = {current_number}")
                                            values.append(current_number)
                                            numeric_count += 1
                                        
                                        if numeric_count == 1:
                                            # Add previous year
                                            print(f"    [DEBUG]   Next div[{j}] (Year {previous_year}): "
                                                  f"'{div_text}' -> Split: '{prev_year_text}' = {prev_number}")
                                            values.append(prev_number)
                                            numeric_count += 1
                                        
                                        # Got both values, break
                                        if numeric_count >= 2:
                                            break
                                        continue
                            
                            # Check if this div looks like it contains a single formatted number
                            # It should be relatively short and contain digits with possible formatting
                            if len(div_text) > 100:
                                # Too long, probably not a single number (unless it's concatenated, which we handled above)
                                continue
                            
                            # Extract number from single number text
                            number = extract_number(div_text)
                            
                            # Validate the number is reasonable (not infinity and not too large)
                            # Indonesian Rupiah values are typically in millions/billions, so cap at 1e15
                            if number > 0 and number < 1e15 and number != float('inf'):
                                year_label = selected_year if numeric_count == 0 else previous_year
                                print(f"    [DEBUG]   Next div[{j}] (Year {year_label}): "
                                      f"'{div_text}' -> {number}")
                                values.append(number)
                                numeric_count += 1
                            elif number == 0 and any(char.isdigit() for char in div_text) and len(div_text) < 50:
                                # Zero value is valid if it's a short text with digits
                                year_label = selected_year if numeric_count == 0 else previous_year
                                print(f"    [DEBUG]   Next div[{j}] (Year {year_label}): "
                                      f"'{div_text}' -> {number}")
                                values.append(number)
                                numeric_count += 1
                except ValueError:
                    print(f"    [DEBUG] Identifier '{identifier}' found but couldn't find its index")
                    return values
                
                if not values:
                    print(f"    [DEBUG] Identifier '{identifier}' found but no numeric values extracted from next divs")
                elif len(values) == 1:
                    print(f"    [DEBUG] Identifier '{identifier}' found but only 1 value extracted")
                
                return values
            
            result = {}
            previous_year = str(int(selected_year) - 1)
            
            # Use provided city and bank, or extract from page source
            print("    [INFO] Extracting Kota/Kabupaten and Bank...")
            extracted_city = city if city else ""
            extracted_bank = bank if bank else ""
            
            # Only try to extract from page if not provided
            if not extracted_city or not extracted_bank:
                try:
                    # Switch to default content to access form fields
                    self.driver.switch_to.default_content()
                    
                    # Try to find city and bank from input fields using Selenium
                    if not extracted_city:
                        try:
                            # Look for city input field
                            city_inputs = [
                                self.driver.find_elements(By.XPATH, "//input[contains(@id, 'City') or contains(@id, 'Kota') or contains(@id, 'Kabupaten')]"),
                                self.driver.find_elements(By.XPATH, "//input[contains(@name, 'City') or contains(@name, 'Kota') or contains(@name, 'Kabupaten')]")
                            ]
                            for inputs in city_inputs:
                                for inp in inputs:
                                    try:
                                        value = inp.get_attribute('value')
                                        if value and value.strip():
                                            extracted_city = value.strip()
                                            print(f"    [DEBUG] Found city from input field: '{extracted_city}'")
                                            break
                                    except:
                                        continue
                                if extracted_city:
                                    break
                        except Exception as e:
                            print(f"    [DEBUG] Could not find city input: {e}")
                    
                    if not extracted_bank:
                        try:
                            # Look for bank input field
                            bank_inputs = [
                                self.driver.find_elements(By.XPATH, "//input[contains(@id, 'Bank')]"),
                                self.driver.find_elements(By.XPATH, "//input[contains(@name, 'Bank')]")
                            ]
                            for inputs in bank_inputs:
                                for inp in inputs:
                                    try:
                                        value = inp.get_attribute('value')
                                        if value and value.strip():
                                            extracted_bank = value.strip()
                                            print(f"    [DEBUG] Found bank from input field: '{extracted_bank}'")
                                            break
                                    except:
                                        continue
                                if extracted_bank:
                                    break
                        except Exception as e:
                            print(f"    [DEBUG] Could not find bank input: {e}")
                    
                    # Fallback: Try to find from BeautifulSoup parsed page
                    if not extracted_city or not extracted_bank:
                        city_inputs = soup.find_all('input', {'id': lambda x: x and ('city' in x.lower() or 'kota' in x.lower() or 'kabupaten' in x.lower())})
                        bank_inputs = soup.find_all('input', {'id': lambda x: x and 'bank' in x.lower()})
                        
                        if not extracted_city:
                            for inp in city_inputs:
                                value = inp.get('value', '')
                                if value and value.strip():
                                    extracted_city = value.strip()
                                    print(f"    [DEBUG] Found city from soup: '{extracted_city}'")
                                    break
                        
                        if not extracted_bank:
                            for inp in bank_inputs:
                                value = inp.get('value', '')
                                if value and value.strip():
                                    extracted_bank = value.strip()
                                    print(f"    [DEBUG] Found bank from soup: '{extracted_bank}'")
                                    break
                
                except Exception as e:
                    print(f"    [WARNING] Could not extract city/bank: {e}")
            
            result['city'] = extracted_city if extracted_city else "N/A"
            result['bank'] = extracted_bank if extracted_bank else "N/A"
            
            # Extract Kredit data only - Sum of 4 identifiers
            print("    [INFO] Extracting Kredit data...")
            # Only extract: Kepada BPR, Kepada Bank Umum, pihak terkait (first occurrence), pihak tidak terkait (first occurrence)
            kredit_identifiers = [
                "Kepada BPR",
                "Kepada Bank Umum",
                "pihak terkait",
                "pihak tidak terkait"
            ]
            kredit_selected_year = 0
            kredit_previous_year = 0
            found_identifiers = set()  # Track which identifiers were found
            
            for identifier in kredit_identifiers:
                values = find_and_extract(identifier)
                if len(values) >= 2:
                    # Check if we already found this identifier (avoid duplicates - only take first occurrence)
                    identifier_key = identifier.strip().lower()
                    if identifier_key not in found_identifiers:
                        # First div = current year, second div = previous year
                        kredit_selected_year += values[0]  # First div (current year)
                        kredit_previous_year += values[1]  # Second div (previous year)
                        found_identifiers.add(identifier_key)
                        print(f"    [DEBUG] Added Kredit from '{identifier}': {values[0]} (2024) + {values[1]} (2023)")
                elif len(values) == 1:
                    identifier_key = identifier.strip().lower()
                    if identifier_key not in found_identifiers:
                        kredit_selected_year += values[0]  # Only current year available
                        found_identifiers.add(identifier_key)
                        print(f"    [DEBUG] Added Kredit from '{identifier}': {values[0]} (2024 only)")
            
            result[f'Kredit {selected_year}'] = kredit_selected_year
            result[f'Kredit {previous_year}'] = kredit_previous_year
            
            print(f"    [OK] Extracted Kredit data: {selected_year}={kredit_selected_year}, {previous_year}={kredit_previous_year}")
            
            # Extract Total Aset data - only one div per year (no sum needed)
            print("    [INFO] Extracting Total Aset data...")
            total_aset_identifier = "Total Aset"
            total_aset_values = find_and_extract(total_aset_identifier)
            
            if len(total_aset_values) >= 2:
                result[f'Total Aset {selected_year}'] = total_aset_values[0]
                result[f'Total Aset {previous_year}'] = total_aset_values[1]
                print(f"    [OK] Extracted Total Aset data: {selected_year}={total_aset_values[0]}, {previous_year}={total_aset_values[1]}")
            elif len(total_aset_values) == 1:
                result[f'Total Aset {selected_year}'] = total_aset_values[0]
                result[f'Total Aset {previous_year}'] = 0
                print(f"    [OK] Extracted Total Aset data: {selected_year}={total_aset_values[0]}, {previous_year}=0 (only current year found)")
            else:
                result[f'Total Aset {selected_year}'] = 0
                result[f'Total Aset {previous_year}'] = 0
                print(f"    [WARNING] Total Aset data not found")
            
            # Extract DPK data - Sum of Tabungan, Deposito, and Simpanan dari Bank Lain for each year
            print("    [INFO] Extracting DPK data (Tabungan + Deposito + Simpanan dari Bank Lain)...")
            dpk_identifiers = [
                "Tabungan",
                "Deposito",
                "Simpanan dari Bank Lain"
            ]
            dpk_selected_year = 0
            dpk_previous_year = 0
            found_dpk_identifiers = set()
            
            for identifier in dpk_identifiers:
                values = find_and_extract(identifier)
                if len(values) >= 2:
                    identifier_key = identifier.strip().lower()
                    if identifier_key not in found_dpk_identifiers:
                        dpk_selected_year += values[0]  # First div (current year)
                        dpk_previous_year += values[1]  # Second div (previous year)
                        found_dpk_identifiers.add(identifier_key)
                        print(f"    [DEBUG] Added DPK from '{identifier}': {values[0]} (2024) + {values[1]} (2023)")
                elif len(values) == 1:
                    identifier_key = identifier.strip().lower()
                    if identifier_key not in found_dpk_identifiers:
                        dpk_selected_year += values[0]  # Only current year available
                        found_dpk_identifiers.add(identifier_key)
                        print(f"    [DEBUG] Added DPK from '{identifier}': {values[0]} (2024 only)")
            
            result[f'DPK {selected_year}'] = dpk_selected_year
            result[f'DPK {previous_year}'] = dpk_previous_year
            
            print(f"    [OK] Extracted DPK data: {selected_year}={dpk_selected_year}, {previous_year}={dpk_previous_year}")
            print(f"    [OK] Total extracted data points: {len(result)}")
            
            # Switch back to default content after extraction
            try:
                self.driver.switch_to.default_content()
                print("    [DEBUG] Switched back to default content")
            except:
                pass
            
            return result
            
        except Exception as e:
            print(f"    [ERROR] Error extracting report data: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def _save_to_excel(self, data: dict, year: str) -> bool:
        """
        Save extracted data to Excel file
        
        Args:
            data: Dictionary with extracted data (must contain: city, bank, Kredit, Total Aset, DPK for both years)
            year: Selected year for filename
            
        Returns:
            True if successful, False otherwise
        """
        if not Workbook:
            print("    [ERROR] openpyxl not installed. Cannot save to Excel.")
            return False
        
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Financial Data"
            
            previous_year = str(int(year) - 1)
            row = 1
            
            # Row 1: Kota/Kabupaten
            city = data.get('city', 'N/A')
            ws[f'A{row}'] = f"Kota/Kabupaten: {city}"
            ws[f'A{row}'].font = Font(bold=True)
            row += 1
            
            # Row 2: Bank
            bank = data.get('bank', 'N/A')
            ws[f'A{row}'] = f"Bank: {bank}"
            ws[f'A{row}'].font = Font(bold=True)
            row += 1
            
            # Row 3: Kredit current year
            kredit_current = data.get(f'Kredit {year}', 0)
            ws[f'A{row}'] = f"Kredit {year}:"
            ws[f'A{row}'].font = Font(bold=True)
            if isinstance(kredit_current, (int, float)):
                ws[f'B{row}'] = f"{kredit_current:,.0f}".replace(',', '.')
            else:
                ws[f'B{row}'] = kredit_current
            row += 1
            
            # Row 4: Kredit previous year
            kredit_previous = data.get(f'Kredit {previous_year}', 0)
            ws[f'A{row}'] = f"Kredit {previous_year}:"
            ws[f'A{row}'].font = Font(bold=True)
            if isinstance(kredit_previous, (int, float)):
                ws[f'B{row}'] = f"{kredit_previous:,.0f}".replace(',', '.')
            else:
                ws[f'B{row}'] = kredit_previous
            row += 1
            
            # Row 5: Total Aset current year
            total_aset_current = data.get(f'Total Aset {year}', 0)
            ws[f'A{row}'] = f"Total Aset {year}:"
            ws[f'A{row}'].font = Font(bold=True)
            if isinstance(total_aset_current, (int, float)):
                ws[f'B{row}'] = f"{total_aset_current:,.0f}".replace(',', '.')
            else:
                ws[f'B{row}'] = total_aset_current
            row += 1
            
            # Row 6: Total Aset previous year
            total_aset_previous = data.get(f'Total Aset {previous_year}', 0)
            ws[f'A{row}'] = f"Total Aset {previous_year}:"
            ws[f'A{row}'].font = Font(bold=True)
            if isinstance(total_aset_previous, (int, float)):
                ws[f'B{row}'] = f"{total_aset_previous:,.0f}".replace(',', '.')
            else:
                ws[f'B{row}'] = total_aset_previous
            row += 1
            
            # Row 7: DPK current year
            dpk_current = data.get(f'DPK {year}', 0)
            ws[f'A{row}'] = f"DPK {year}:"
            ws[f'A{row}'].font = Font(bold=True)
            if isinstance(dpk_current, (int, float)):
                ws[f'B{row}'] = f"{dpk_current:,.0f}".replace(',', '.')
            else:
                ws[f'B{row}'] = dpk_current
            row += 1
            
            # Row 8: DPK previous year
            dpk_previous = data.get(f'DPK {previous_year}', 0)
            ws[f'A{row}'] = f"DPK {previous_year}:"
            ws[f'A{row}'].font = Font(bold=True)
            if isinstance(dpk_previous, (int, float)):
                ws[f'B{row}'] = f"{dpk_previous:,.0f}".replace(',', '.')
            else:
                ws[f'B{row}'] = dpk_previous
            
            # Auto-adjust column widths
            ws.column_dimensions['A'].width = 60
            ws.column_dimensions['B'].width = 20
            
            # Save file
            filename = f"OJK_Report_{year}.xlsx"
            filepath = self.output_dir / filename
            wb.save(filepath)
            print(f"    [OK] Data saved to {filepath}")
            print(f"    [OK] Excel content:")
            print(f"      - Kota/Kabupaten: {city}")
            print(f"      - Bank: {bank}")
            print(f"      - Kredit {year}: {kredit_current:,.0f}".replace(',', '.'))
            print(f"      - Kredit {previous_year}: {kredit_previous:,.0f}".replace(',', '.'))
            print(f"      - Total Aset {year}: {total_aset_current:,.0f}".replace(',', '.'))
            print(f"      - Total Aset {previous_year}: {total_aset_previous:,.0f}".replace(',', '.'))
            print(f"      - DPK {year}: {dpk_current:,.0f}".replace(',', '.'))
            print(f"      - DPK {previous_year}: {dpk_previous:,.0f}".replace(',', '.'))
            print("\n  [INFO] Excel file successfully saved. Closing browser...")
            self.cleanup()
            return True
            
        except Exception as e:
            print(f"    [ERROR] Error saving to Excel: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def _save_to_csv(self, filepath: Path, data: list):
        """
        Save data to CSV file
        
        Args:
            filepath: Path to save CSV file
            data: List of dictionaries representing rows
        """
        if not data:
            return
        
        # Get all unique keys from all rows
        fieldnames = set()
        for row in data:
            fieldnames.update(row.keys())
        fieldnames = sorted(list(fieldnames))
        
        with open(filepath, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(data)
    
    def cleanup(self, kill_processes: bool = False):
        """
        Close browser and cleanup all Selenium resources
        
        Args:
            kill_processes: If True, also kill any lingering Chrome/ChromeDriver processes
        """
        print("[INFO] Membersihkan sumber daya Selenium...")
        
        # Cleanup ExtJS helper first
        if self.extjs:
            try:
                self.extjs = None
            except:
                pass
        
        # Cleanup WebDriverWait
        if self.wait:
            try:
                self.wait = None
            except:
                pass
        
        # Cleanup WebDriver - this is the most important
        if self.driver:
            try:
                # Close all windows
                try:
                    self.driver.close()
                except:
                    pass
                
                # Quit driver (closes all windows and ends session)
                self.driver.quit()
                print("[OK] Browser ditutup")
            except Exception as e:
                print(f"[WARNING] Error saat menutup browser: {e}")
            finally:
                # Clear reference
                self.driver = None
        
        # Force garbage collection to free memory
        import gc
        gc.collect()
        print("[OK] Sumber daya Selenium telah dibersihkan")
        
        # Optionally kill lingering processes
        if kill_processes:
            try:
                from .utils import kill_chrome_processes
                kill_chrome_processes()
            except ImportError:
                try:
                    from utils import kill_chrome_processes
                    kill_chrome_processes()
                except ImportError:
                    print("[WARNING] Tidak dapat mengimpor fungsi kill_chrome_processes")
    
    def unload_selenium(self, kill_processes: bool = True):
        """
        Explicitly unload Selenium and optionally kill lingering processes
        
        Args:
            kill_processes: If True, kill any lingering Chrome/ChromeDriver processes (default: True)
        """
        self.cleanup(kill_processes=kill_processes)
    
    def __enter__(self):
        """Context manager entry"""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit"""
        self.cleanup()


if __name__ == "__main__":
    # Example usage
    scraper = OJKExtJSScraper(headless=False)
    try:
        scraper.initialize()
        scraper.navigate_to_page()
        scraper.select_tab_bpr_konvensional()
        scraper.scrape_all_data(month="Desember", year="2024")
    finally:
        scraper.cleanup()

