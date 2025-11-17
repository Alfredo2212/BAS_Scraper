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
        
        # Test: Set month, year, and select province
        print("\n[TEST] Testing month, year, and province selection...")
        scraper.scrape_all_data(month="Desember", year="2024")
        
    except Exception as e:
        print(f"\n[ERROR] Test failed: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # Cleanup Selenium resources
        print("\n[OK] Test completed")
        print("[INFO] Membersihkan sumber daya Selenium...")
        # Use kill_processes=True to ensure all Chrome processes are terminated
        scraper.cleanup(kill_processes=True)
        print("[OK] Semua sumber daya telah dibersihkan")


if __name__ == "__main__":
    main()

