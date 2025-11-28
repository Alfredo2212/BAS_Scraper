"""
Test script for Sindikasi bank finder
Tests finding banks in both BPR Konvensional and BPR Syariah URLs
"""

import sys
import importlib.util
from pathlib import Path

# Import using importlib to handle directory name with spaces
module_path = Path(__file__).parent / "Laporan Publikasi Sindikasi" / "scraper.py"
spec = importlib.util.spec_from_file_location("sindikasi_scraper", module_path)
sindikasi_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(sindikasi_module)
SindikasiScraper = sindikasi_module.SindikasiScraper

def test_bank_finder():
    """Test the bank finder with a sample list file"""
    print("Testing Sindikasi bank finder...")
    print("=" * 70)
    
    # Create scraper instance
    scraper = SindikasiScraper(headless=False)  # Set to False to see browser
    
    # Test with the existing list file
    list_file = Path("input/list_28_11_2025")
    
    if not list_file.exists():
        print(f"[ERROR] List file not found: {list_file}")
        return
    
    try:
        # Run the bank finder
        scraper.find_all_banks(list_file)
        print("\n" + "=" * 70)
        print("Bank finder test completed!")
    except KeyboardInterrupt:
        print("\n[INFO] Test interrupted by user")
        scraper.cleanup()
    except Exception as e:
        print(f"\n[ERROR] Test failed: {e}")
        import traceback
        traceback.print_exc()
        scraper.cleanup()

if __name__ == "__main__":
    test_bank_finder()

