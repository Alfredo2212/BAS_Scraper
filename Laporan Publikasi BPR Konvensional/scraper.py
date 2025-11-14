"""
OJK ExtJS Scraper
Main scraping logic using ExtJS API exclusively
No DOM clicking, pure ExtJS ComponentQuery
"""

import time
import csv
from pathlib import Path
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

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
        
        # Wait for page to fully load (reduced by 80%)
        print("[INFO] Waiting for page to fully load...")
        time.sleep(0.6)  # Reduced from 3s to 0.6s (80% faster)
        
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
        self._select_dropdowns_and_checkboxes()
        
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
                        
                        # Wait for grid to load
                        if not self.extjs.wait_for_grid(timeout=3):  # Reduced from 15s to 3s (80% faster)
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
    
    def _select_dropdowns_and_checkboxes(self):
        """
        3-step process:
        1. Click dropdown arrow ext-gen1064 and select topmost <li>
        2. Click dropdown arrow ext-gen1069 and select topmost <li>
        3. Find treeview elements and check checkboxes inside them
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
        
        # Step 3: Find treeview elements and check checkboxes
        print("\n  [Step 4.3] Finding treeview elements and checking checkboxes...")
        time.sleep(0.3)  # Wait a bit longer for treeview to be ready
        
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
                
                # Debug: Print all nested elements to understand structure
                all_divs = treeview_element.find_elements(By.XPATH, ".//div")
                print(f"    [DEBUG] Found {len(all_divs)} div elements inside {treeview_id}")
                
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
                            checkbox_class = checkbox.get_attribute("class")
                            
                            print(f"    [DEBUG] Checkbox {i+1} - role: {role}, aria-checked: {aria_checked}, type: {checkbox_type}, class: {checkbox_class}")
                            
                            # Check if checkbox is already checked
                            if aria_checked == "true" or (checkbox_type == "checkbox" and checkbox.is_selected()):
                                print(f"    [INFO] Checkbox {i+1} already checked in {treeview_id}")
                                continue
                            
                            # Scroll to checkbox and click
                            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", checkbox)
                            time.sleep(0.1)  # Wait for scroll
                            
                            # Try clicking with JavaScript first
                            try:
                                self.driver.execute_script("arguments[0].click();", checkbox)
                                print(f"    [OK] Checked checkbox {i+1} in {treeview_id} (JavaScript click)")
                            except:
                                # Fallback to regular click
                                checkbox.click()
                                print(f"    [OK] Checked checkbox {i+1} in {treeview_id} (regular click)")
                            
                            time.sleep(0.1)  # Reduced by 80%
                        except Exception as e:
                            print(f"    [WARNING] Could not check checkbox {i+1} in {treeview_id}: {e}")
                            import traceback
                            traceback.print_exc()
                else:
                    # Debug: Print HTML structure to help understand what's inside
                    print(f"    [WARNING] No checkboxes found in {treeview_id}")
                    try:
                        inner_html = treeview_element.get_attribute("innerHTML")
                        if inner_html:
                            print(f"    [DEBUG] Inner HTML preview (first 500 chars): {inner_html[:500]}")
                    except:
                        pass
            except Exception as e:
                print(f"    [WARNING] Could not find treeview element {treeview_id}: {e}")
                import traceback
                traceback.print_exc()
        
        print("  [OK] Completed 3-step dropdown and checkbox selection")
        
        # Step 4: Click "Tampilkan" button after checkboxes are checked
        print("\n  [Step 4.4] Clicking 'Tampilkan' button...")
        time.sleep(0.2)  # Reduced by 80%
        
        try:
            # Find the Tampilkan button by ID
            tampilkan_button = self.driver.find_element(By.ID, "ShowReportButton-btnInnerEl")
            print("  [OK] Found 'Tampilkan' button")
            
            # Scroll to button and click
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", tampilkan_button)
            time.sleep(0.1)  # Wait for scroll
            
            # Try clicking with JavaScript first
            try:
                self.driver.execute_script("arguments[0].click();", tampilkan_button)
                print("  [OK] Clicked 'Tampilkan' button (JavaScript click)")
            except:
                # Fallback to regular click
                tampilkan_button.click()
                print("  [OK] Clicked 'Tampilkan' button (regular click)")
            
            time.sleep(0.3)  # Wait for form submission/PostBack
        except Exception as e:
            print(f"  [WARNING] Could not click 'Tampilkan' button: {e}")
            # Try alternative: find by text content
            try:
                tampilkan_button = self.driver.find_element(By.XPATH, "//span[contains(text(), 'Tampilkan')]")
                print("  [OK] Found 'Tampilkan' button by text")
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", tampilkan_button)
                time.sleep(0.1)
                self.driver.execute_script("arguments[0].click();", tampilkan_button)
                print("  [OK] Clicked 'Tampilkan' button (alternative method)")
                time.sleep(0.3)
            except Exception as e2:
                print(f"  [ERROR] Could not click 'Tampilkan' button with alternative method: {e2}")
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

