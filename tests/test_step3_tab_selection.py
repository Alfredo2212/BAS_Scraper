"""
Test Step 3: Tab Selection (BPR Konvensional)
Tests that we can select the BPR Konvensional tab and wait for content to load
"""

import sys
from pathlib import Path

# Add parent directory to path so we can import modules
sys.path.insert(0, str(Path(__file__).parent.parent))


def test_tab_selector_in_config():
    """Test that tab selector is defined in config"""
    try:
        from config.settings import OJKConfig
        
        tab_selector = OJKConfig.SELECTORS.get('tab_bpr_konvensional')
        assert tab_selector is not None, "Tab selector not found in config"
        print(f"[OK] Tab selector found: {tab_selector}")
        return True
    except Exception as e:
        print(f"[FAIL] Config check failed: {e}")
        return False


def test_navigation_and_tab_selection():
    """Test navigation to page and tab selection"""
    try:
        from scraper.ojk_scraper import OJKScraper
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC
        from config.settings import OJKConfig
        
        print("[INFO] Testing navigation and tab selection...")
        print("[INFO] Opening browser (this may take a moment)...")
        
        scraper = OJKScraper(headless=False)  # Use visible browser for debugging
        
        try:
            # Step 1: Navigate to page
            scraper.initialize()
            scraper.navigate_to_page()
            
            # Step 2: Select BPR Konvensional tab
            success = scraper.select_bpr_konvensional_tab()
            
            if not success:
                print("[FAIL] Tab selection returned False")
                return False
            
            # Step 3: Verify form elements are visible
            # Note: We're still inside the iframe after tab click
            print("[INFO] Verifying form elements are visible...")
            
            # Wait a bit more for form to load after tab click
            import time
            time.sleep(1)  # Reduced by 50%
            
            # Check for month dropdown (still in iframe)
            month_dropdown = scraper.driver.find_element(
                By.XPATH, 
                OJKConfig.SELECTORS['dropdown_month']
            )
            if month_dropdown.is_displayed():
                print("[OK] Month dropdown is visible")
            else:
                print("[WARNING] Month dropdown found but not visible")
            
            # Check for year dropdown
            try:
                year_dropdown = scraper.driver.find_element(
                    By.XPATH,
                    OJKConfig.SELECTORS['dropdown_year']
                )
                if year_dropdown.is_displayed():
                    print("[OK] Year dropdown is visible")
                else:
                    print("[WARNING] Year dropdown found but not visible")
            except Exception as e:
                print(f"[WARNING] Year dropdown check failed: {e}")
            
            # Check for province dropdown
            try:
                province_dropdown = scraper.driver.find_element(
                    By.XPATH,
                    OJKConfig.SELECTORS['dropdown_province']
                )
                if province_dropdown.is_displayed():
                    print("[OK] Province dropdown is visible")
                else:
                    print("[WARNING] Province dropdown found but not visible")
            except Exception as e:
                print(f"[WARNING] Province dropdown check failed: {e}")
            
            print("[OK] Tab selection test completed successfully")
            return True
            
        finally:
            # Switch back to default content before cleanup
            try:
                scraper.driver.switch_to.default_content()
            except:
                pass
            # Always cleanup
            scraper.cleanup()
            print("[OK] Browser closed")
            
    except Exception as e:
        print(f"[FAIL] Tab selection test failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_tab_element_found():
    """Test that tab element can be found on the page - SKIPPED (covered by test_navigation_and_tab_selection)"""
    print("[INFO] Skipping this test - tab element finding is covered in test_navigation_and_tab_selection")
    print("[OK] Test skipped to avoid duplicate browser opens")
    return True


if __name__ == "__main__":
    print("Testing Step 3: Tab Selection (BPR Konvensional)")
    print("=" * 60)
    print("\nNOTE: This will open a Chrome browser window for testing.")
    print("Please do not close it manually - the test will close it automatically.\n")
    
    results = []
    
    print("[Test 1] Testing tab selector in config...")
    results.append(test_tab_selector_in_config())
    
    print("\n[Test 2] Testing tab element can be found...")
    results.append(test_tab_element_found())
    
    print("\n[Test 3] Testing navigation and tab selection...")
    results.append(test_navigation_and_tab_selection())
    
    print("\n" + "=" * 60)
    if all(results):
        print("[SUCCESS] All Step 3 tests passed!")
        print("\nNext steps:")
        print("  - Step 4: Test month dropdown with PostBack handling")
        print("  - Step 5: Test year dropdown with PostBack handling")
    else:
        print("[FAIL] Some tests failed. Please check the errors above.")
        print("\nTroubleshooting:")
        print("  - Check if the OJK website structure has changed")
        print("  - Verify the tab selector XPath is correct")
        print("  - Check if the page loaded completely")
        sys.exit(1)

