"""
OJK Scraper - Main scraper class
Orchestrates the scraping process using modular components
"""

from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support.ui import WebDriverWait

from config.settings import OJKConfig
from scraper.selenium_setup import SeleniumSetup
from scraper.postback_handler import PostBackHandler
from scraper.data_extractor import DataExtractor
from scraper.excel_exporter import ExcelExporter


class OJKScraper:
    """Main scraper class for OJK publication reports"""
    
    def __init__(self, headless: bool = None):
        """
        Initialize the scraper
        
        Args:
            headless: Whether to run browser in headless mode. If None, uses OJKConfig.HEADLESS_MODE
        """
        self.base_url = OJKConfig.BASE_URL
        self.driver: WebDriver = None
        self.wait: WebDriverWait = None
        self.postback_handler: PostBackHandler = None
        self.data_extractor: DataExtractor = None
        self.headless = headless if headless is not None else OJKConfig.HEADLESS_MODE
    
    def initialize(self):
        """Initialize WebDriver and helper classes"""
        if self.driver is None:
            self.driver = SeleniumSetup.create_driver(headless=self.headless)
            self.wait = SeleniumSetup.create_wait(self.driver)
            self.postback_handler = PostBackHandler(self.driver, self.wait)
            self.data_extractor = DataExtractor(self.driver, self.wait)
    
    def navigate_to_page(self):
        """Navigate to OJK publication reports page"""
        if self.driver is None:
            self.initialize()
        
        print(f"[INFO] Navigating to: {self.base_url}")
        self.driver.get(self.base_url)
        
        # Wait for page to be fully loaded with multiple checks
        import time
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.common.by import By
        
        print("[INFO] Waiting for page to fully load...")
        
        # Wait for document ready state
        try:
            WebDriverWait(self.driver, 10).until(
                lambda d: d.execute_script('return document.readyState') == 'complete'
            )
            print("[OK] Document ready state: complete")
        except Exception as e:
            print(f"[WARNING] Ready state check timeout: {e}")
        
        # Wait for jQuery to finish (if present)
        try:
            WebDriverWait(self.driver, 5).until(
                lambda d: d.execute_script('return typeof jQuery === "undefined" || jQuery.active === 0')
            )
            print("[OK] jQuery activity completed")
        except:
            pass  # jQuery might not be present
        
        # Wait for any tab elements to appear (indicates page structure is loaded)
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//span[contains(@class, 'x-tab') or contains(@class, 'x-tab-inner')]"))
            )
            print("[OK] Tab elements detected on page")
        except Exception as e:
            print(f"[WARNING] Tab elements not found yet: {e}")
        
        # Additional wait for page to stabilize
        print("[INFO] Waiting for page to stabilize...")
        time.sleep(1)  # Reduced by 50%
        
        print("[OK] Page loaded successfully")
    
    def select_bpr_konvensional_tab(self):
        """
        Select the 'BPR Konvensional' tab and wait for content to load
        
        Returns:
            bool: True if tab was selected successfully, False otherwise
        """
        if self.driver is None:
            self.initialize()
        
        if self.wait is None:
            self.wait = SeleniumSetup.create_wait(self.driver)
        
        try:
            from selenium.webdriver.common.by import By
            from selenium.webdriver.support import expected_conditions as EC
            from selenium.webdriver.support.ui import WebDriverWait
            from selenium.common.exceptions import TimeoutException, WebDriverException
            import time
            
            print("[INFO] Waiting for page to be ready before looking for tab...")
            
            # Wait for page to be fully loaded first
            try:
                WebDriverWait(self.driver, 15).until(
                    lambda d: d.execute_script('return document.readyState') == 'complete'
                )
            except:
                pass  # Continue even if ready state check fails
            
            # Switch to iframe first (tabs are inside iframe)
            print("[INFO] Looking for iframe containing the tabs...")
            iframe_switched = False
            try:
                # Switch back to default content first (in case we're already in an iframe)
                self.driver.switch_to.default_content()
                
                # Find all iframes
                iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
                print(f"[DEBUG] Found {len(iframes)} iframe(s) on page")
                
                if iframes:
                    # Try to switch to each iframe and check if it contains tabs
                    for i, iframe in enumerate(iframes):
                        try:
                            print(f"[INFO] Trying to switch to iframe {i+1}...")
                            self.driver.switch_to.frame(iframe)
                            
                            # Wait a moment for iframe content to load
                            time.sleep(1)  # Reduced by 50%
                            
                            # Check if this iframe contains tab elements
                            try:
                                WebDriverWait(self.driver, 3).until(  # Reduced by 50%
                                    EC.presence_of_element_located((By.XPATH, "//span[contains(@class, 'x-tab-inner')]"))
                                )
                                print(f"[OK] Switched to iframe {i+1} - tabs found inside!")
                                iframe_switched = True
                                break
                            except TimeoutException:
                                # This iframe doesn't have tabs, try next one
                                print(f"[DEBUG] Iframe {i+1} doesn't contain tabs, trying next...")
                                self.driver.switch_to.default_content()
                                continue
                        except Exception as iframe_error:
                            print(f"[WARNING] Could not switch to iframe {i+1}: {iframe_error}")
                            self.driver.switch_to.default_content()
                            continue
                    
                    if not iframe_switched:
                        print("[WARNING] Could not find iframe with tabs, trying first iframe anyway...")
                        self.driver.switch_to.default_content()
                        self.driver.switch_to.frame(iframes[0])
                        time.sleep(1.5)  # Reduced by 50%
                        iframe_switched = True
                else:
                    print("[INFO] No iframes found - tabs should be in main page")
            except Exception as e:
                print(f"[WARNING] Error handling iframes: {e}")
                # Try to continue anyway
            
            # Wait for tab container to be present (indicates tabs are loaded)
            print("[INFO] Waiting for tab container to load...")
            # Try waiting longer for dynamic content
            max_attempts = 5
            tabs_found = False
            for attempt in range(max_attempts):
                try:
                    WebDriverWait(self.driver, 3).until(
                        EC.presence_of_element_located((By.XPATH, "//span[contains(@class, 'x-tab-inner')]"))
                    )
                    print("[OK] Tab container detected")
                    tabs_found = True
                    break
                except TimeoutException:
                    print(f"[DEBUG] Attempt {attempt + 1}/{max_attempts}: Tabs not found yet, waiting...")
                    time.sleep(1)  # Reduced by 50%
            
            if tabs_found:
                # DEBUG: Find all tab elements
                all_tabs = self.driver.find_elements(By.XPATH, "//span[contains(@class, 'x-tab-inner')]")
                print(f"[DEBUG] Found {len(all_tabs)} tab elements on page")
                for i, tab in enumerate(all_tabs[:5]):  # Show first 5
                    try:
                        print(f"[DEBUG] Tab {i+1}: text='{tab.text}', class='{tab.get_attribute('class')}'")
                    except:
                        pass
            else:
                print("[WARNING] Tab container not found after multiple attempts")
                # DEBUG: Try to find any tab-like elements
                try:
                    any_tabs = self.driver.find_elements(By.XPATH, "//*[contains(@class, 'tab') or contains(@class, 'x-tab')]")
                    print(f"[DEBUG] Found {len(any_tabs)} elements with 'tab' in class name")
                    
                    # DEBUG: Check page title and URL
                    print(f"[DEBUG] Current page title: {self.driver.title}")
                    print(f"[DEBUG] Current URL: {self.driver.current_url}")
                    
                    # DEBUG: Look for any elements with 'BPR' text
                    bpr_elements = self.driver.find_elements(By.XPATH, "//*[contains(text(), 'BPR')]")
                    print(f"[DEBUG] Found {len(bpr_elements)} elements containing 'BPR' text")
                    for i, elem in enumerate(bpr_elements[:5]):
                        try:
                            print(f"[DEBUG] BPR Element {i+1}: tag={elem.tag_name}, text='{elem.text[:50]}', class='{elem.get_attribute('class')}'")
                        except:
                            pass
                except Exception as debug_error:
                    print(f"[DEBUG] Error during debug: {debug_error}")
            
            # Additional wait for dynamic content
            time.sleep(0.5)  # Reduced by 50%
            
            print("[INFO] Looking for 'BPR Konvensional' tab...")
            print(f"[DEBUG] Using XPath selector: {OJKConfig.SELECTORS['tab_bpr_konvensional']}")
            
            # Try multiple selector strategies with reasonable timeout
            tab_xpath = OJKConfig.SELECTORS['tab_bpr_konvensional']
            
            # Use 10 second timeout (should be enough for most cases)
            extended_wait = WebDriverWait(self.driver, 10)
            
            # Wait for tab to be present first (more lenient)
            print("[INFO] Waiting for tab element to appear...")
            tab_element = None
            try:
                tab_element = extended_wait.until(
                    EC.presence_of_element_located((By.XPATH, tab_xpath))
                )
                print("[OK] Tab element found with primary selector")
            except TimeoutException as e:
                # Try alternative: find by text first, then get parent button
                print(f"[DEBUG] Primary selector failed: {e}")
                print("[INFO] Primary selector failed, trying alternative approach...")
                try:
                    # Find the inner span with text (more specific)
                    inner_xpath = "//span[contains(@class, 'x-tab-inner') and contains(text(), 'BPR Konvensional')]"
                    print(f"[DEBUG] Trying alternative XPath: {inner_xpath}")
                    inner_span = extended_wait.until(
                        EC.presence_of_element_located((By.XPATH, inner_xpath))
                    )
                    print("[OK] Found inner span with text")
                    print(f"[DEBUG] Inner span text: '{inner_span.text}', tag: '{inner_span.tag_name}'")
                    print(f"[DEBUG] Inner span HTML: {inner_span.get_attribute('outerHTML')[:200]}")
                    
                    # Get the parent button element or anchor
                    try:
                        tab_element = inner_span.find_element(By.XPATH, "./ancestor::span[contains(@class, 'x-tab-button')]")
                        print("[OK] Found parent span with x-tab-button class")
                    except:
                        try:
                            # Try finding anchor parent
                            tab_element = inner_span.find_element(By.XPATH, "./ancestor::a[contains(@class, 'x-tab')]")
                            print("[OK] Found parent anchor with x-tab class")
                        except:
                            # Try any ancestor that might be clickable
                            tab_element = inner_span.find_element(By.XPATH, "./ancestor::*[contains(@class, 'x-tab')][1]")
                            print("[OK] Found ancestor element with x-tab class")
                except Exception as alt_error:
                    print(f"[ERROR] Alternative selector also failed: {alt_error}")
                    # DEBUG: Try to find element by partial text match
                    print("[DEBUG] Attempting to find element by partial text match...")
                    try:
                        all_elements = self.driver.find_elements(By.XPATH, "//*[contains(text(), 'BPR')]")
                        print(f"[DEBUG] Found {len(all_elements)} elements containing 'BPR'")
                        for elem in all_elements[:3]:
                            try:
                                print(f"[DEBUG] Element: tag={elem.tag_name}, text='{elem.text[:50]}', class='{elem.get_attribute('class')}'")
                            except:
                                pass
                    except:
                        pass
                    raise
            
            if tab_element is None:
                raise Exception("Could not find tab element with any selector")
            
            # DEBUG: Print element details
            print(f"[DEBUG] Tab element found:")
            print(f"[DEBUG]   Tag name: {tab_element.tag_name}")
            print(f"[DEBUG]   Text: '{tab_element.text}'")
            print(f"[DEBUG]   Class: '{tab_element.get_attribute('class')}'")
            print(f"[DEBUG]   ID: '{tab_element.get_attribute('id')}'")
            print(f"[DEBUG]   Is displayed: {tab_element.is_displayed()}")
            print(f"[DEBUG]   Is enabled: {tab_element.is_enabled()}")
            print(f"[DEBUG]   Location: {tab_element.location}")
            print(f"[DEBUG]   Size: {tab_element.size}")
            
            # Scroll into view to ensure element is visible
            print("[INFO] Scrolling tab into view...")
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", tab_element)
            time.sleep(0.5)  # Reduced by 50%
            
            # DEBUG: Check element state after scroll
            print(f"[DEBUG] After scroll - Is displayed: {tab_element.is_displayed()}, Is enabled: {tab_element.is_enabled()}")
            
            # Wait for element to be clickable
            print("[INFO] Waiting for tab to be clickable...")
            try:
                extended_wait.until(EC.element_to_be_clickable(tab_element))
                print("[OK] Element is clickable")
            except TimeoutException:
                print("[WARNING] Element not clickable yet, but will try to click anyway")
                print(f"[DEBUG] Element state - displayed: {tab_element.is_displayed()}, enabled: {tab_element.is_enabled()}")
            
            print(f"[INFO] Tab found: '{tab_element.text if tab_element.text else 'BPR Konvensional'}'")
            print("[INFO] Clicking 'BPR Konvensional' tab...")
            
            # Save screenshot before clicking (for debugging)
            try:
                screenshot_path = "logs/debug_before_click.png"
                self.driver.save_screenshot(screenshot_path)
                print(f"[DEBUG] Screenshot saved to {screenshot_path}")
            except:
                pass
            
            # Try JavaScript click first (more reliable for dynamic content)
            try:
                print("[DEBUG] Attempting JavaScript click...")
                self.driver.execute_script("arguments[0].click();", tab_element)
                print("[OK] JavaScript click executed")
            except Exception as js_error:
                print(f"[WARNING] JavaScript click failed: {js_error}")
                print(f"[DEBUG] Error details: {type(js_error).__name__}: {str(js_error)}")
                try:
                    print("[DEBUG] Attempting regular click...")
                    tab_element.click()
                    print("[OK] Regular click executed")
                except WebDriverException as click_error:
                    print(f"[ERROR] Both click methods failed")
                    print(f"[DEBUG] Regular click error: {type(click_error).__name__}: {str(click_error)}")
                    # DEBUG: Try to get more info about why click failed
                    try:
                        print(f"[DEBUG] Element still displayed: {tab_element.is_displayed()}")
                        print(f"[DEBUG] Element still enabled: {tab_element.is_enabled()}")
                        # Try ActionChains as last resort
                        from selenium.webdriver.common.action_chains import ActionChains
                        print("[DEBUG] Trying ActionChains click...")
                        ActionChains(self.driver).move_to_element(tab_element).click().perform()
                        print("[OK] ActionChains click executed")
                    except Exception as ac_error:
                        print(f"[ERROR] ActionChains also failed: {ac_error}")
                        raise
            
            # Wait for tab content to load (PostBack or AJAX)
            print("[INFO] Waiting for tab content to load (PostBack may take time)...")
            time.sleep(1.5)  # Reduced by 50%: Wait for PostBack to start
            
            # Wait for jQuery/AJAX to finish (if present)
            try:
                WebDriverWait(self.driver, 5).until(  # Reduced by 50%
                    lambda d: d.execute_script('return typeof jQuery === "undefined" || jQuery.active === 0')
                )
                print("[OK] jQuery/AJAX activity completed")
            except:
                pass
            
            # Wait for document ready state in iframe
            try:
                WebDriverWait(self.driver, 5).until(  # Reduced by 50%
                    lambda d: d.execute_script('return document.readyState') == 'complete'
                )
            except:
                pass
            
            # DEBUG: Check what elements are present after tab click
            print("[DEBUG] Checking what elements are present after tab click...")
            try:
                all_selects = self.driver.find_elements(By.TAG_NAME, "select")
                print(f"[DEBUG] Found {len(all_selects)} select elements in iframe")
                for i, sel in enumerate(all_selects[:5]):
                    try:
                        sel_id = sel.get_attribute('id') or 'no-id'
                        print(f"[DEBUG] Select {i+1}: id='{sel_id}'")
                    except:
                        pass
            except:
                pass
            
            # Verify that form elements are now visible (check for month dropdown)
            # This confirms the tab content has loaded
            month_dropdown_xpath = OJKConfig.SELECTORS['dropdown_month']
            print("[INFO] Verifying form elements are loaded...")
            try:
                extended_wait.until(
                    EC.presence_of_element_located((By.XPATH, month_dropdown_xpath))
                )
                print("[OK] Tab selected and content loaded successfully")
                return True
            except TimeoutException:
                print("[WARNING] Tab clicked but form elements not found - page may still be loading")
                print("[INFO] Waiting additional 1.5 seconds for PostBack...")
                time.sleep(1.5)  # Reduced by 50%
                # Try one more time
                try:
                    self.driver.find_element(By.XPATH, month_dropdown_xpath)
                    print("[OK] Form elements found after additional wait")
                    return True
                except:
                    print("[WARNING] Form elements still not found")
                    # DEBUG: Try to find any select elements
                    try:
                        selects = self.driver.find_elements(By.XPATH, "//select")
                        print(f"[DEBUG] Found {len(selects)} select elements total")
                        if selects:
                            print("[INFO] Select elements exist but may have different IDs")
                            print("[OK] Tab was clicked successfully - form may need different selector")
                            return True
                    except:
                        pass
                    # Return True anyway, as the tab was clicked
                    return True
            
        except Exception as e:
            print(f"[ERROR] Failed to select tab: {e}")
            import traceback
            traceback.print_exc()
            # Switch back to default content on error
            try:
                self.driver.switch_to.default_content()
            except:
                pass
            return False
        # Note: We stay in the iframe after successful tab click
        # because form elements are also inside the iframe
        # Caller should switch back to default_content when done
        
    def __enter__(self):
        """Context manager entry"""
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit - cleanup"""
        self.cleanup()
    
    def cleanup(self):
        """Clean up resources"""
        if self.driver:
            self.driver.quit()
            self.driver = None
            self.wait = None
            self.postback_handler = None
            self.data_extractor = None

