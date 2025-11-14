"""
Test Step 2: Selenium WebDriver setup and page navigation
Tests that WebDriver can be initialized and navigate to OJK website
"""

import sys
from pathlib import Path

# Add parent directory to path so we can import modules
sys.path.insert(0, str(Path(__file__).parent.parent))


def test_selenium_setup_import():
    """Test that SeleniumSetup can be imported"""
    try:
        from scraper.selenium_setup import SeleniumSetup
        print("[OK] SeleniumSetup imported successfully")
        return True
    except ImportError as e:
        print(f"[FAIL] SeleniumSetup import failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_webdriver_creation():
    """Test that WebDriver can be created"""
    try:
        from scraper.selenium_setup import SeleniumSetup
        
        print("[INFO] Creating WebDriver (this may take a moment)...")
        driver = SeleniumSetup.create_driver(headless=False)  # Use visible browser for debugging
        
        if driver is None:
            print("[FAIL] WebDriver is None")
            return False
        
        print("[OK] WebDriver created successfully")
        print(f"  Browser: {driver.name}")
        print(f"  Window size: {driver.get_window_size()}")
        
        # Cleanup
        driver.quit()
        print("[OK] WebDriver closed successfully")
        return True
        
    except Exception as e:
        print(f"[FAIL] WebDriver creation failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_page_navigation():
    """Test navigation to OJK website"""
    try:
        from scraper.ojk_scraper import OJKScraper
        from config.settings import OJKConfig
        
        print(f"[INFO] Testing navigation to: {OJKConfig.BASE_URL}")
        print("[INFO] Opening browser (this may take a moment)...")
        
        scraper = OJKScraper(headless=False)  # Use visible browser for debugging
        
        try:
            # Initialize and navigate
            scraper.initialize()
            scraper.navigate_to_page()
            
            # Verify page loaded
            page_title = scraper.driver.title
            current_url = scraper.driver.current_url
            
            print(f"[OK] Page loaded successfully")
            print(f"  Page title: {page_title}")
            print(f"  Current URL: {current_url}")
            
            # Check if we're on the right page
            if "ojk.go.id" in current_url.lower():
                print("[OK] Successfully navigated to OJK website")
                return True
            else:
                print(f"[WARNING] URL doesn't contain 'ojk.go.id': {current_url}")
                return False
                
        finally:
            # Always cleanup
            scraper.cleanup()
            print("[OK] Browser closed")
            
    except Exception as e:
        print(f"[FAIL] Navigation test failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_wait_creation():
    """Test that WebDriverWait can be created"""
    try:
        from scraper.selenium_setup import SeleniumSetup
        
        print("[INFO] Testing WebDriverWait creation...")
        driver = SeleniumSetup.create_driver(headless=False)
        
        try:
            wait = SeleniumSetup.create_wait(driver)
            
            if wait is None:
                print("[FAIL] WebDriverWait is None")
                return False
            
            print(f"[OK] WebDriverWait created successfully")
            print(f"  Timeout: {wait._timeout} seconds")
            return True
            
        finally:
            driver.quit()
            
    except Exception as e:
        print(f"[FAIL] WebDriverWait creation failed: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    print("Testing Step 2: Selenium WebDriver Setup & Navigation")
    print("=" * 60)
    print("\nNOTE: This will open a Chrome browser window for testing.")
    print("Please do not close it manually - the test will close it automatically.\n")
    
    results = []
    
    print("[Test 1] Testing SeleniumSetup import...")
    results.append(test_selenium_setup_import())
    
    print("\n[Test 2] Testing WebDriver creation...")
    results.append(test_webdriver_creation())
    
    print("\n[Test 3] Testing WebDriverWait creation...")
    results.append(test_wait_creation())
    
    print("\n[Test 4] Testing page navigation...")
    results.append(test_page_navigation())
    
    print("\n" + "=" * 60)
    if all(results):
        print("[SUCCESS] All Step 2 tests passed!")
        print("\nNext steps:")
        print("  - Step 3: Test tab selection (BPR Konvensional)")
        print("  - Step 4: Test dropdown interactions with PostBack")
    else:
        print("[FAIL] Some tests failed. Please check the errors above.")
        sys.exit(1)

