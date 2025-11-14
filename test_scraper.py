"""
Simple test script for OJK ExtJS Scraper
Tests the scraper with direct URL approach
"""

import sys
from pathlib import Path

# Add the parent directory to path to import the module
sys.path.insert(0, str(Path(__file__).parent))

# Import using importlib to handle directory name with spaces
import importlib.util
module_path = Path(__file__).parent / "Laporan Publikasi BPR Konvensional" / "scraper.py"
spec = importlib.util.spec_from_file_location("scraper_module", module_path)
scraper_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(scraper_module)
OJKExtJSScraper = scraper_module.OJKExtJSScraper


def main():
    """Test the ExtJS scraper"""
    print("OJK ExtJS Scraper Test")
    print("=" * 60)
    print("This will navigate directly to BPR Konvensional report page")
    print("and test ExtJS combobox interactions\n")
    
    scraper = OJKExtJSScraper(headless=False)
    try:
        scraper.initialize()
        scraper.navigate_to_page()
        scraper.select_tab_bpr_konvensional()
        
        # Test: Set month, year, and select province
        print("\n[TEST] Testing month, year, and province selection...")
        scraper.scrape_all_data(month="Desember", year="2024")
        
    except Exception as e:
        print(f"\n[ERROR] Test failed: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # Don't close browser - keep it open for inspection
        print("\n[OK] Test completed")
        print("[INFO] Browser will remain open for inspection. Close it manually when done.")
        print("[INFO] Press ENTER to exit and close the browser...")
        try:
            input()  # Wait for user to press Enter before script ends
        except EOFError:
            # Handle non-interactive environments (like CI/CD)
            print("[INFO] Non-interactive environment detected. Browser will remain open.")
            import time
            time.sleep(300)  # Wait 5 minutes as fallback
        # Optionally close browser when user presses Enter:
        # scraper.cleanup()


if __name__ == "__main__":
    main()

