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
        time.sleep(3)
        
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
            time.sleep(1)
        
        # If not in main page, check for iframes
        print("[INFO] ExtJS not in main page, checking for iframes...")
        iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
        if iframes:
            print(f"[INFO] Found {len(iframes)} iframe(s), checking inside...")
            for i, iframe in enumerate(iframes):
                try:
                    self.driver.switch_to.frame(iframe)
                    print(f"[INFO] Switched to iframe {i+1}")
                    time.sleep(0.4)  # Reduced from 2s to 0.4s (80% faster)
                    
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
        time.sleep(0.6)  # Reduced from 3s to 0.6s (80% faster)
        
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
    
    def select_tab_bpr_konvensional(self):
        """
        No-op: We're already on BPR Konvensional page via direct URL
        
        Returns:
            True (always succeeds since we're already on the right page)
        """
        print("[INFO] Already on BPR Konvensional report page (direct URL)")
        return True
    
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
        time.sleep(0.18)  # Reduced to 30% (0.6s * 0.3 = 0.18s)
        
        # Try to find and click the trigger arrow with ID ext-gen1050 (static ID for month dropdown)
        print("[INFO] Looking for trigger arrow (id='ext-gen1050')...")
        trigger_found = False
        max_attempts = 10
        wait_interval = 0.4  # Reduced from 2s to 0.4s (80% faster)
        
        for attempt in range(max_attempts):
            try:
                # Try to find by ID
                trigger = self.driver.find_element(By.ID, "ext-gen1050")
                print(f"[OK] Found trigger arrow (attempt {attempt + 1})")
                
                # Click the trigger to open dropdown
                print("[INFO] Clicking trigger arrow to open month dropdown...")
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", trigger)
                time.sleep(0.05)  # Reduced from 0.25s to 0.05s (80% faster)
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
            time.sleep(0.2)  # Reduced from 1s to 0.2s (80% faster)
            
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
                    time.sleep(0.05)  # Reduced from 0.25s to 0.05s (80% faster)
                    self.driver.execute_script("arguments[0].click();", target_li)
                    print(f"[OK] Clicked <li> element with text '{month}'")
                    time.sleep(0.2)  # Reduced from 1s to 0.2s (80% faster) - Wait for PostBack
                else:
                    available_options = [li.text.strip() for li in li_elements if li.text.strip()]
                    print(f"[WARNING] Could not find <li> with text '{month}'. Available: {available_options}")
            except Exception as e:
                print(f"[WARNING] Could not click <li> element: {e}")
        
        # Optimize: Start ticking checkboxes immediately after month selection (multitasking)
        print("\n[OPTIMIZATION] Starting checkbox selection in parallel...")
        self._tick_checkboxes()
        
        # Now try ExtJS API method for other operations
        print("[INFO] Checking for ExtJS availability...")
        max_attempts = 10
        for attempt in range(max_attempts):
            if self.extjs.check_extjs_available():
                print("[OK] ExtJS is available")
                break
            time.sleep(0.2)  # Reduced from 1s to 0.2s (80% faster)
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
            time.sleep(0.6)  # Reduced from 3s to 0.6s (80% faster) - Wait a bit more
            combos = self.extjs.list_all_combos()
            if combos:
                print(f"[OK] Found {len(combos)} comboboxes after additional wait")
            else:
                print("[WARNING] No comboboxes found - will try to continue anyway")
        
        # Step 2: Select year by typing directly into input field
        print(f"\n[Step 2] Selecting year: 2024 (hardcoded)")
        time.sleep(0.2)  # Reduced from 1s to 0.2s (80% faster)
        
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
            time.sleep(0.2)  # Reduced from 1s to 0.2s (80% faster) - Wait for PostBack
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
        time.sleep(0.2)  # Reduced from 1s to 0.2s (80% faster)
        
        # Try to find and click the trigger arrow with ID ext-gen1059 (static ID for province dropdown)
        print("[INFO] Looking for province trigger arrow (id='ext-gen1059')...")
        province_trigger_found = False
        max_attempts = 10
        wait_interval = 0.4  # Reduced from 2s to 0.4s (80% faster)
        
        for attempt in range(max_attempts):
            try:
                # Try to find by ID
                province_trigger = self.driver.find_element(By.ID, "ext-gen1059")
                print(f"[OK] Found province trigger arrow (attempt {attempt + 1})")
                
                # Click the trigger to open dropdown
                print("[INFO] Clicking province trigger arrow to open dropdown...")
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", province_trigger)
                time.sleep(0.05)  # Reduced from 0.25s to 0.05s (80% faster)
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
            time.sleep(0.2)  # Reduced from 1s to 0.2s (80% faster)
            
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
                    time.sleep(0.05)  # Reduced from 0.25s to 0.05s (80% faster)
                    self.driver.execute_script("arguments[0].click();", target_li)
                    print(f"[OK] Clicked <li> element with text '{province_name}'")
                    time.sleep(0.2)  # Reduced from 1s to 0.2s (80% faster) - Wait for PostBack
                else:
                    available_options = [li.text.strip() for li in li_elements if li.text.strip()]
                    print(f"[WARNING] Could not find <li> with text '{province_name}'. Available: {available_options[:10]}...")
            except Exception as e:
                print(f"[WARNING] Could not click province <li> element: {e}")
        
        # Step 4: Select dropdowns and check checkboxes (3-step process)
        print("\n[Step 4] Starting 3-step dropdown and checkbox selection...")
        self._select_dropdowns_and_checkboxes(year)
        
        # Get all provinces using ExtJS API (for later use in loop)
        print("\n[Step 3b] Getting all provinces via ExtJS...")
        province_combo_name = self._find_combo_name_by_keyword("province")
        if not province_combo_name:
            province_combo_name = self.extjs.find_combo_by_position(2)
        
        provinces = self.extjs.get_extjs_combo_values(province_combo_name) if province_combo_name else []
        print(f"[INFO] Found {len(provinces)} provinces: {provinces[:5]}...")
        
        # Step 4: Loop through provinces
        for province in provinces:
            if not province or province.strip() == '':
                continue
            
            print(f"\n{'='*60}")
            print(f"[PROVINCE] Processing: {province}")
            print(f"{'='*60}")
            
            # Select province
            self.extjs.set_extjs_combo(province_combo_name, province)
            time.sleep(0.2)  # Reduced from 1s to 0.2s (80% faster) - Wait for PostBack to load cities
            
            # Get all cities for this province
            city_combo_name = self._find_combo_name_by_keyword("city")
            if not city_combo_name:
                city_combo_name = self.extjs.find_combo_by_position(3)
            
            cities = self.extjs.get_extjs_combo_values(city_combo_name) if city_combo_name else []
            print(f"[INFO] Found {len(cities)} cities in {province}")
            
            # Step 5: Loop through cities
            for city in cities:
                if not city or city.strip() == '':
                    continue
                
                print(f"\n  [CITY] Processing: {city}")
                
                # Select city
                self.extjs.set_extjs_combo(city_combo_name, city)
                time.sleep(0.2)  # Reduced from 1s to 0.2s (80% faster) - Wait for PostBack to load banks
                
                # Get all banks for this city
                bank_combo_name = self._find_combo_name_by_keyword("bank")
                if not bank_combo_name:
                    bank_combo_name = self.extjs.find_combo_by_position(4)
                
                banks = self.extjs.get_extjs_combo_values(bank_combo_name) if bank_combo_name else []
                print(f"    [INFO] Found {len(banks)} banks in {city}")
                
                # Step 6: Loop through banks
                for bank in banks:
                    if not bank or bank.strip() == '':
                        continue
                    
                    print(f"\n    [BANK] Processing: {bank}")
                    
                    try:
                        # Select bank
                        self.extjs.set_extjs_combo(bank_combo_name, bank)
                        time.sleep(0.1)  # Reduced from 0.5s to 0.1s (80% faster)
                        
                        # Click "Tampilkan" button
                        if not self.extjs.click_tampilkan():
                            print(f"      [WARNING] Failed to click Tampilkan for {bank}")
                            continue
                        
                        # Wait for grid to load (10 seconds as required for Tampilkan to load)
                        print(f"      [INFO] Waiting up to 7.5 seconds for grid to load after Tampilkan...")
                        if not self.extjs.wait_for_grid(timeout=7.5):  # Increased to 7.5s to ensure proper loading
                            print(f"      [WARNING] Grid not loaded for {bank}")
                            continue
                        
                        # Extract grid data
                        grid_data = self.extjs.get_grid_data()
                        
                        if grid_data:
                            # Save to CSV
                            filename = f"ojk_{province}_{city}_{bank}_{month}_{year}.csv"
                            filename = filename.replace('/', '_').replace('\\', '_')  # Sanitize filename
                            filepath = self.output_dir / filename
                            
                            if grid_data:
                                self._save_to_csv(filepath, grid_data)
                                print(f"      [OK] Saved {len(grid_data)} rows to {filepath}")
                            else:
                                print(f"      [WARNING] No data found for {bank}")
                        else:
                            print(f"      [WARNING] Could not extract grid data for {bank}")
                        
                        time.sleep(0.2)  # Reduced from 1s to 0.2s (80% faster) - Delay between requests
                        
                    except Exception as e:
                        print(f"      [ERROR] Error processing {bank}: {e}")
                        continue
        
        print("\n[OK] Scraping completed!")
    
    def _tick_checkboxes(self):
        """
        Find treeview elements and check the first 3 checkboxes
        This can be done in parallel with other operations for faster execution
        """
        print("  [INFO] Finding treeview elements and checking checkboxes...")
        time.sleep(0.3)  # Wait a bit for treeview to be ready
        
        treeview_ids = [
            "treeview-1012-record-BPK-901-000001",
            "treeview-1012-record-BPK-901-000002",
            "treeview-1012-record-BPK-901-000003"
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
                            time.sleep(0.03)  # Reduced to 30% (0.1s * 0.3 = 0.03s) - Wait for scroll
                            
                            # Try clicking with JavaScript first
                            try:
                                self.driver.execute_script("arguments[0].click();", checkbox)
                                print(f"    [OK] Checked checkbox {i+1} in {treeview_id} (JavaScript click)")
                            except:
                                # Fallback to regular click
                                checkbox.click()
                                print(f"    [OK] Checked checkbox {i+1} in {treeview_id} (regular click)")
                            
                            time.sleep(0.03)  # Reduced to 30% (0.1s * 0.3 = 0.03s) - Wait after click
                        except Exception as e:
                            print(f"    [WARNING] Could not check checkbox {i+1} in {treeview_id}: {e}")
                else:
                    print(f"    [WARNING] No checkboxes found in {treeview_id}")
            except Exception as e:
                print(f"    [WARNING] Could not find treeview element {treeview_id}: {e}")
        
        print("  [OK] Completed checkbox selection")
    
    def _select_dropdowns_and_checkboxes(self, year: str):
        """
        2-step process (checkboxes are now ticked earlier for optimization):
        1. Click dropdown arrow ext-gen1064 and select topmost <li>
        2. Click dropdown arrow ext-gen1069 and select topmost <li>
        3. Click Tampilkan button and extract data
        
        Args:
            year: Selected year for data extraction
        """
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        
        # Step 1: Click dropdown arrow ext-gen1064 and select topmost <li>
        print("\n  [Step 4.1] Clicking dropdown arrow ext-gen1064...")
        time.sleep(0.2)  # Reduced by 80%
        
        try:
            dropdown1_trigger = self.driver.find_element(By.ID, "ext-gen1064")
            print("  [OK] Found dropdown arrow ext-gen1064")
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown1_trigger)
            time.sleep(0.05)  # Reduced by 80%
            self.driver.execute_script("arguments[0].click();", dropdown1_trigger)
            print("  [OK] Clicked dropdown arrow ext-gen1064")
            
            # Wait for dropdown and select topmost <li>
            time.sleep(0.2)  # Reduced by 80%
            wait = WebDriverWait(self.driver, 5)
            dropdown_list = wait.until(
                EC.presence_of_element_located((By.XPATH, "//ul[contains(@class, 'x-list-plain')]"))
            )
            
            li_elements = self.driver.find_elements(By.XPATH, "//li[@role='option' or contains(@class, 'x-boundlist-item')]")
            if not li_elements:
                li_elements = self.driver.find_elements(By.XPATH, "//ul[contains(@class, 'x-list-plain')]//li")
            
            # Filter out empty <li> elements and get the first non-empty one
            for li in li_elements:
                try:
                    li_text = li.text.strip()
                    if li_text:  # Only select non-empty <li>
                        print(f"  [OK] Selecting topmost <li>: '{li_text}'")
                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", li)
                        time.sleep(0.05)  # Reduced by 80%
                        self.driver.execute_script("arguments[0].click();", li)
                        print(f"  [OK] Clicked topmost <li>: '{li_text}'")
                        time.sleep(0.2)  # Reduced by 80% - Wait for PostBack
                        break
                except:
                    continue
        except Exception as e:
            print(f"  [WARNING] Could not select dropdown ext-gen1064: {e}")
        
        # Step 2: Click dropdown arrow ext-gen1069 and select topmost <tr>
        print("\n  [Step 4.2] Clicking dropdown arrow ext-gen1069...")
        time.sleep(0.2)  # Reduced by 80%
        
        try:
            dropdown2_trigger = self.driver.find_element(By.ID, "ext-gen1069")
            print("  [OK] Found dropdown arrow ext-gen1069")
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown2_trigger)
            time.sleep(0.05)  # Reduced by 80%
            self.driver.execute_script("arguments[0].click();", dropdown2_trigger)
            print("  [OK] Clicked dropdown arrow ext-gen1069")
            
            # Wait for dropdown to appear and find the dropdown container
            time.sleep(0.3)  # Wait a bit longer for dropdown to appear
            wait = WebDriverWait(self.driver, 10)
            
            # Try to find the dropdown container (boundlist or table)
            dropdown_container = None
            try:
                # Look for boundlist container first
                dropdown_container = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'x-boundlist') or contains(@class, 'x-layer')]"))
                )
                print("  [OK] Found dropdown container (boundlist/layer)")
            except:
                # Try table container
                try:
                    dropdown_container = wait.until(
                        EC.presence_of_element_located((By.XPATH, "//table[contains(@class, 'x-boundlist') or ancestor::div[contains(@class, 'x-boundlist')]]"))
                    )
                    print("  [OK] Found dropdown container (table)")
                except:
                    print("  [WARNING] Could not find specific dropdown container, searching globally")
            
            # Find all <tr> elements, scoped to dropdown container if found
            if dropdown_container:
                tr_elements = dropdown_container.find_elements(By.XPATH, ".//tr")
            else:
                # Search globally but prioritize visible dropdown elements
                tr_elements = self.driver.find_elements(By.XPATH, "//div[contains(@class, 'x-boundlist')]//tr | //div[contains(@class, 'x-layer')]//tr")
                if not tr_elements:
                    tr_elements = self.driver.find_elements(By.XPATH, "//tr")
            
            if tr_elements:
                print(f"  [DEBUG] Found {len(tr_elements)} <tr> elements")
                # Filter out empty <tr> elements and get the first non-empty one
                for i, tr in enumerate(tr_elements):
                    try:
                        # Check if element is visible
                        if not tr.is_displayed():
                            continue
                            
                        tr_text = tr.text.strip()
                        if tr_text:  # Only select non-empty <tr>
                            print(f"  [OK] Selecting topmost <tr> (index {i}): '{tr_text[:50]}...'")
                            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", tr)
                            time.sleep(0.1)  # Slightly longer wait for visibility
                            self.driver.execute_script("arguments[0].click();", tr)
                            print(f"  [OK] Clicked topmost <tr>: '{tr_text[:50]}...'")
                            time.sleep(0.2)  # Reduced by 80% - Wait for PostBack
                            break
                    except Exception as e:
                        print(f"  [DEBUG] Could not click tr at index {i}: {e}")
                        continue
                else:
                    print("  [WARNING] No clickable <tr> elements found")
            else:
                # Fallback: if no <tr> found, try <li> as before
                print("  [INFO] No <tr> elements found, trying <li> as fallback...")
                if dropdown_container:
                    li_elements = dropdown_container.find_elements(By.XPATH, ".//li[@role='option' or contains(@class, 'x-boundlist-item')]")
                else:
                    li_elements = self.driver.find_elements(By.XPATH, "//li[@role='option' or contains(@class, 'x-boundlist-item')]")
                if not li_elements:
                    li_elements = self.driver.find_elements(By.XPATH, "//ul[contains(@class, 'x-list-plain')]//li")
                
                for li in li_elements:
                    try:
                        if not li.is_displayed():
                            continue
                        li_text = li.text.strip()
                        if li_text:  # Only select non-empty <li>
                            print(f"  [OK] Selecting topmost <li> (fallback): '{li_text}'")
                            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", li)
                            time.sleep(0.05)  # Reduced by 80%
                            self.driver.execute_script("arguments[0].click();", li)
                            print(f"  [OK] Clicked topmost <li> (fallback): '{li_text}'")
                            time.sleep(0.2)  # Reduced by 80% - Wait for PostBack
                            break
                    except:
                        continue
        except Exception as e:
            print(f"  [WARNING] Could not select dropdown ext-gen1069: {e}")
            import traceback
            traceback.print_exc()
        
        print("  [OK] Completed dropdown selection (checkboxes already ticked earlier)")
        
        # Step 4: Click "Tampilkan" button and wait for report to load
        print("\n  [Step 4.4] Clicking 'Tampilkan' button and waiting for report to load...")
        time.sleep(0.2)  # Reduced by 80%
        
        # Click Tampilkan button once and wait 1 minute
        try:
            # Find and click Tampilkan button
            tampilkan_button = None
            try:
                tampilkan_button = self.driver.find_element(By.ID, "ShowReportButton-btnInnerEl")
                print("  [OK] Found 'Tampilkan' button")
            except:
                # Try alternative: find by text content
                tampilkan_button = self.driver.find_element(By.XPATH, "//span[contains(text(), 'Tampilkan')]")
                print("  [OK] Found 'Tampilkan' button by text")
            
            # Scroll to button and click
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", tampilkan_button)
            time.sleep(0.1)  # Wait for scroll
            
            # Try clicking with JavaScript first
            try:
                self.driver.execute_script("arguments[0].click();", tampilkan_button)
                print("  [OK] Clicked 'Tampilkan' button")
            except:
                # Fallback to regular click
                tampilkan_button.click()
                print("  [OK] Clicked 'Tampilkan' button (regular click)")
            
            # Wait 1 minute and check for identifiers
            print("  [INFO] Waiting 1 minute for report to load and checking for identifiers...")
            self._wait_for_report_loaded(max_wait=60)
            print("  [OK] Wait completed. Proceeding to extract data...")
            
        except Exception as e:
            print(f"  [WARNING] Error clicking 'Tampilkan' button: {e}")
            # Still wait 1 minute even if button click failed
            print("  [INFO] Waiting 1 minute anyway...")
            self._wait_for_report_loaded(max_wait=60)
        
        # Always proceed to extract data (will create Excel with whatever is found)
        report_loaded = True
        
        if report_loaded:
            # Extract data from the generated report
            print("\n  [Step 4.5] Extracting data from report...")
            extracted_data = self._extract_report_data(year)
            
            if extracted_data:
                # Save to Excel
                print("\n  [Step 4.6] Saving data to Excel...")
                success = self._save_to_excel(extracted_data, year)
                if success:
                    return  # Browser will be closed in _save_to_excel
            else:
                print("  [ERROR] Failed to extract data from report")
        else:
            print("  [ERROR] Report did not load. Cannot extract data.")
    
    def _find_combo_name_by_keyword(self, keyword: str) -> str:
        """
        Find combobox name by keyword
        
        Args:
            keyword: Keyword to search for (e.g., "month", "year", "province")
            
        Returns:
            Component name if found, empty string otherwise
        """
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
        check_interval = 1  # Check every 1 second
        
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
    
    def _extract_report_data(self, selected_year: str) -> dict:
        """
        Extract financial data from the generated report
        
        Args:
            selected_year: The selected year (e.g., "2024")
            
        Returns:
            Dictionary with extracted data
        """
        try:
            # Try to find report in iframe first, then main page
            print("    [DEBUG] Checking for report in iframes...")
            iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
            page_source = None
            
            # Check iframes for report content
            for iframe in iframes:
                try:
                    self.driver.switch_to.frame(iframe)
                    iframe_source = self.driver.page_source
                    if "Kredit" in iframe_source or "Aset" in iframe_source or "DPK" in iframe_source:
                        print("    [DEBUG] Found report content in iframe")
                        page_source = iframe_source
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
            else:
                # Switch back to default content after getting iframe source
                self.driver.switch_to.default_content()
            
            soup = BeautifulSoup(page_source, 'html.parser')
            
            # Find all div elements (data is in divs, not tds)
            all_divs = soup.find_all('div')
            print(f"    [DEBUG] Found {len(all_divs)} div elements in page")
            
            # Helper function to find numeric value from text
            def extract_number(text):
                """Extract numeric value from text, removing commas (thousand separators) for proper parsing"""
                if not text:
                    return 0
                original_text = text
                # Remove HTML entities and whitespace
                text = text.replace('&nbsp;', '').strip()
                # Remove commas (thousand separators) - e.g., "1,234,567" -> "1234567"
                text = text.replace(',', '')
                # Remove all non-numeric characters except decimal point
                cleaned_text = re.sub(r'[^\d.]', '', text)
                try:
                    result = float(cleaned_text) if cleaned_text else 0
                    if original_text.strip() != cleaned_text:
                        print(f"      [DEBUG]     Raw: '{original_text}' -> Cleaned: '{cleaned_text}' -> Number: {result}")
                    return result
                except:
                    print(f"      [DEBUG]     Failed to extract number from: '{original_text}'")
                    return 0
            
            # Helper function to find identifier and get next 2 div values
            def find_and_extract(identifier):
                """Find identifier text and extract next 2 numeric values from sibling div elements"""
                values = []
                found = False
                
                # Find all divs in the page
                all_divs = soup.find_all('div')
                
                # Search for the identifier div
                identifier_div = None
                identifier_index = -1
                
                for idx, div in enumerate(all_divs):
                    div_text = div.get_text(strip=True).replace('&nbsp;', ' ')
                    if identifier.lower() in div_text.lower():
                        print(f"    [DEBUG] Found identifier '{identifier}' in div[{idx}]: '{div_text[:60]}'")
                        identifier_div = div
                        identifier_index = idx
                        found = True
                        break
                
                if found and identifier_div:
                    # Find the parent container (likely a row or container with multiple divs)
                    parent = identifier_div.parent
                    
                    # Get all sibling divs in the same parent
                    if parent:
                        sibling_divs = parent.find_all('div', recursive=False)
                        
                        # Find the index of our identifier div
                        identifier_sibling_index = -1
                        for idx, sibling in enumerate(sibling_divs):
                            if sibling == identifier_div:
                                identifier_sibling_index = idx
                                break
                        
                        # Get the next 2 sibling divs after the identifier
                        if identifier_sibling_index >= 0:
                            next_siblings = sibling_divs[identifier_sibling_index + 1:identifier_sibling_index + 3]
                            
                            if len(next_siblings) >= 2:
                                for j, sibling_div in enumerate(next_siblings[:2]):
                                    div_text = sibling_div.get_text(strip=True)
                                    year_label = selected_year if j == 0 else previous_year
                                    print(f"    [DEBUG]   Next sibling div[{j+1}] (Year {year_label}) raw HTML: {str(sibling_div)[:150]}")
                                    print(f"    [DEBUG]   Next sibling div[{j+1}] text content: '{div_text}'")
                                    value = extract_number(div_text)
                                    print(f"    [DEBUG]   Next sibling div[{j+1}] -> extracted value: {value} (Year {year_label})")
                                    values.append(value)
                    
                    # If not found in siblings, try finding next divs in document order
                    if len(values) < 2 and identifier_index >= 0:
                        # Get next divs after the identifier div
                        candidate_divs = all_divs[identifier_index + 1:identifier_index + 10]
                        
                        for div in candidate_divs:
                            div_text = div.get_text(strip=True)
                            # Check if div contains a number (likely a data div)
                            if re.search(r'\d', div_text) and len(div_text) < 100:
                                # Check if it's a sibling or in the same parent structure
                                if div.parent == identifier_div.parent or div.parent == parent:
                                    values.append(extract_number(div_text))
                                    print(f"    [DEBUG]   Found data div in document order: '{div_text}' -> {values[-1]}")
                                    if len(values) >= 2:
                                        break
                
                if not found:
                    print(f"    [DEBUG] Identifier '{identifier}' NOT FOUND in page")
                elif not values:
                    print(f"    [DEBUG] Identifier '{identifier}' found but no div values extracted")
                elif len(values) < 2:
                    print(f"    [DEBUG] Identifier '{identifier}' found but only {len(values)} value(s) extracted")
                
                return values
            
            result = {}
            previous_year = str(int(selected_year) - 1)
            
            # Extract Kota/Kabupaten and Bank from page source
            print("    [INFO] Extracting Kota/Kabupaten and Bank...")
            city = ""
            bank = ""
            
            try:
                # Switch to default content to access form fields
                self.driver.switch_to.default_content()
                
                # Try to find city and bank from input fields using Selenium
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
                                    city = value.strip()
                                    print(f"    [DEBUG] Found city from input field: '{city}'")
                                    break
                            except:
                                continue
                        if city:
                            break
                except Exception as e:
                    print(f"    [DEBUG] Could not find city input: {e}")
                
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
                                    bank = value.strip()
                                    print(f"    [DEBUG] Found bank from input field: '{bank}'")
                                    break
                            except:
                                continue
                        if bank:
                            break
                except Exception as e:
                    print(f"    [DEBUG] Could not find bank input: {e}")
                
                # Fallback: Try to find from BeautifulSoup parsed page
                if not city or not bank:
                    city_inputs = soup.find_all('input', {'id': lambda x: x and ('city' in x.lower() or 'kota' in x.lower() or 'kabupaten' in x.lower())})
                    bank_inputs = soup.find_all('input', {'id': lambda x: x and 'bank' in x.lower()})
                    
                    for inp in city_inputs:
                        value = inp.get('value', '')
                        if value and value.strip():
                            city = value.strip()
                            print(f"    [DEBUG] Found city from soup: '{city}'")
                            break
                    
                    for inp in bank_inputs:
                        value = inp.get('value', '')
                        if value and value.strip():
                            bank = value.strip()
                            print(f"    [DEBUG] Found bank from soup: '{bank}'")
                            break
                
            except Exception as e:
                print(f"    [WARNING] Could not extract city/bank: {e}")
            
            result['city'] = city if city else "N/A"
            result['bank'] = bank if bank else "N/A"
            
            # Extract Kredit data only - Sum of 4 identifiers
            print("    [INFO] Extracting Kredit data...")
            kredit_identifiers = [
                "a. Kepada BPR",
                "b. Kepada Bank Umum",
                "c. Kepada non bank – pihak terkait",
                "d. Kepada non bank – pihak tidak terkait"
            ]
            kredit_selected_year = 0
            kredit_previous_year = 0
            found_identifiers = set()  # Track which identifiers were found
            
            for identifier in kredit_identifiers:
                values = find_and_extract(identifier)
                if len(values) >= 2:
                    # Check if we already found this identifier (avoid duplicates)
                    identifier_key = identifier.split('.')[-1].strip() if '.' in identifier else identifier
                    if identifier_key not in found_identifiers:
                        # First div = current year, second div = previous year
                        kredit_selected_year += values[0]  # First div (current year)
                        kredit_previous_year += values[1]  # Second div (previous year)
                        found_identifiers.add(identifier_key)
                        print(f"    [DEBUG] Added Kredit from '{identifier}': {values[0]} (2024) + {values[1]} (2023)")
                elif len(values) == 1:
                    identifier_key = identifier.split('.')[-1].strip() if '.' in identifier else identifier
                    if identifier_key not in found_identifiers:
                        kredit_selected_year += values[0]  # Only current year available
                        found_identifiers.add(identifier_key)
                        print(f"    [DEBUG] Added Kredit from '{identifier}': {values[0]} (2024 only)")
            
            result[f'Kredit {selected_year}'] = kredit_selected_year
            result[f'Kredit {previous_year}'] = kredit_previous_year
            
            print(f"    [OK] Extracted Kredit data: {selected_year}={kredit_selected_year}, {previous_year}={kredit_previous_year}")
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
            data: Dictionary with extracted data (must contain: city, bank, Kredit {year}, Kredit {previous_year})
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
            
            # Row 1: Kota/Kabupaten
            city = data.get('city', 'N/A')
            ws['A1'] = f"Kota/Kabupaten: {city}"
            ws['A1'].font = Font(bold=True)
            
            # Row 2: Bank
            bank = data.get('bank', 'N/A')
            ws['A2'] = f"Bank: {bank}"
            ws['A2'].font = Font(bold=True)
            
            # Row 3: Kredit current year
            kredit_current = data.get(f'Kredit {year}', 0)
            ws['A3'] = f"Kredit {year}:"
            ws['A3'].font = Font(bold=True)
            if isinstance(kredit_current, (int, float)):
                ws['B3'] = f"{kredit_current:,.0f}".replace(',', '.')
            else:
                ws['B3'] = kredit_current
            
            # Row 4: Kredit previous year
            kredit_previous = data.get(f'Kredit {previous_year}', 0)
            ws['A4'] = f"Kredit {previous_year}:"
            ws['A4'].font = Font(bold=True)
            if isinstance(kredit_previous, (int, float)):
                ws['B4'] = f"{kredit_previous:,.0f}".replace(',', '.')
            else:
                ws['B4'] = kredit_previous
            
            # Auto-adjust column widths
            ws.column_dimensions['A'].width = 60
            ws.column_dimensions['B'].width = 20
            
            # Save file
            filename = f"OJK_Report_{year}.xlsx"
            filepath = self.output_dir / filename
            wb.save(filepath)
            print(f"    [OK] Data saved to {filepath}")
            print(f"    [OK] Excel content: Kota/Kabupaten={city}, Bank={bank}, Kredit {year}={kredit_current}, Kredit {previous_year}={kredit_previous}")
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
    
    def cleanup(self):
        """Close browser and cleanup"""
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
    
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

