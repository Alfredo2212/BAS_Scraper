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
        
        # Test: Run all 3 phases sequentially (new 3-phase implementation)
        print("\n[TEST] Testing 3-phase scraping with auto-detection of month/year...")
        print("[TEST] This will run Phase 001, Phase 002, and Phase 003 sequentially")
        print("[TEST] Each phase will have its own Chrome session\n")
        scraper.run_all_phases()  # Will auto-detect month and year, run all 3 phases
        
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

