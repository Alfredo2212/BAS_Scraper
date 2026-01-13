"""
Simple test script for OJK ExtJS Scraper
Tests the scraper with direct URL approach
Saves output to production OSS folder: D:\APP\OSS\client\assets\publikasi
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
    # Production output directory from environment variable
    from config.settings import Settings
    production_output_dir = Settings.OUTPUT_PUBLIKASI
    
    print("OJK ExtJS Scraper Test")
    print("=" * 60)
    print("This will navigate directly to BPR Konvensional report page")
    print("and test ExtJS combobox interactions")
    print(f"Output will be saved to: {production_output_dir}\n")
    
    # Create output directory if it doesn't exist
    try:
        production_output_dir.mkdir(parents=True, exist_ok=True)
        print(f"[OK] Output directory ready: {production_output_dir}\n")
    except Exception as e:
        print(f"[ERROR] Failed to create output directory: {e}")
        return
    
    scraper = OJKExtJSScraper(headless=False)
    try:
        scraper.initialize()
        scraper.navigate_to_page()
        
        # Override output directory to production folder
        scraper.output_dir = production_output_dir
        scraper.output_dir.mkdir(parents=True, exist_ok=True)
        print(f"[INFO] Output directory set to: {production_output_dir}\n")
        
        # Test: Run all 3 phases sequentially (new 3-phase implementation)
        print("\n[TEST] Testing 3-phase scraping with auto-detection of month/year...")
        print("[TEST] This will run Phase 001, Phase 002, and Phase 003 sequentially")
        print("[TEST] Each phase will have its own Chrome session\n")
        scraper.run_all_phases()  # Will auto-detect month and year, run all 3 phases
        
        print(f"\n[OK] Test completed successfully!")
        print(f"[INFO] Output saved to: {production_output_dir}")
        
    except Exception as e:
        print(f"\n[ERROR] Test failed: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # Cleanup Selenium resources
        print("\n[INFO] Membersihkan sumber daya Selenium...")
        # Use kill_processes=True to ensure all Chrome processes are terminated
        scraper.cleanup(kill_processes=True)
        print("[OK] Semua sumber daya telah dibersihkan")


if __name__ == "__main__":
    main()

