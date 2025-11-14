"""
Test Step 1: Verify project setup and dependencies
Tests that all modules can be imported and basic initialization works
"""

import sys
from pathlib import Path

# Add parent directory to path so we can import modules
sys.path.insert(0, str(Path(__file__).parent.parent))


def test_config_imports():
    """Test that config module can be imported"""
    try:
        from config.settings import OJKConfig, Settings
        print("[OK] Config module imported successfully")
        assert OJKConfig.BASE_URL is not None
        assert Settings.OUTPUT_DIR is not None
        return True
    except ImportError as e:
        print(f"[FAIL] Config import failed: {e}")
        return False


def test_scraper_imports():
    """Test that all scraper modules can be imported"""
    try:
        from scraper.ojk_scraper import OJKScraper
        from scraper.selenium_setup import SeleniumSetup
        from scraper.postback_handler import PostBackHandler
        from scraper.data_extractor import DataExtractor
        from scraper.excel_exporter import ExcelExporter
        print("[OK] All scraper modules imported successfully")
        return True
    except ImportError as e:
        print(f"[FAIL] Scraper import failed: {e}")
        return False


def test_scraper_initialization():
    """Test scraper can be initialized"""
    try:
        from scraper.ojk_scraper import OJKScraper
        scraper = OJKScraper()
        assert scraper.base_url is not None
        assert scraper.driver is None  # Should not be initialized yet
        print("[OK] Scraper initialized successfully")
        print(f"  Base URL: {scraper.base_url}")
        return True
    except Exception as e:
        print(f"[FAIL] Initialization failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_directories_exist():
    """Test that required directories exist"""
    try:
        from config.settings import Settings
        assert Settings.OUTPUT_DIR.exists(), f"Output directory does not exist: {Settings.OUTPUT_DIR}"
        print(f"[OK] Output directory exists: {Settings.OUTPUT_DIR}")
        return True
    except Exception as e:
        print(f"[FAIL] Directory check failed: {e}")
        return False


if __name__ == "__main__":
    print("Testing Step 1: Project Setup & Structure")
    print("=" * 60)
    
    results = []
    
    print("\n[Test 1] Testing config module imports...")
    results.append(test_config_imports())
    
    print("\n[Test 2] Testing scraper module imports...")
    results.append(test_scraper_imports())
    
    print("\n[Test 3] Testing scraper initialization...")
    results.append(test_scraper_initialization())
    
    print("\n[Test 4] Testing directory structure...")
    results.append(test_directories_exist())
    
    print("\n" + "=" * 60)
    if all(results):
        print("[SUCCESS] All Step 1 tests passed!")
        print("\nNext steps:")
        print("  - Step 2: Test Selenium WebDriver setup")
        print("  - Step 3: Test page navigation")
    else:
        print("[FAIL] Some tests failed. Please check the errors above.")
        sys.exit(1)

