"""
Test script for Sindikasi list file parser
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

def test_list_parser():
    """Test the list file parser"""
    print("Testing Sindikasi list file parser...")
    print("=" * 70)
    
    # Create scraper instance (no browser needed for parsing)
    scraper = SindikasiScraper(headless=True)
    
    # Test with the existing list file
    list_file = Path("input/list_28_11_2025")
    
    if not list_file.exists():
        print(f"[ERROR] List file not found: {list_file}")
        return
    
    # Read and parse the file
    result = scraper.read_list_file(list_file)
    
    # Display results
    print("\nParsing Results:")
    print(f"  SCRAPE: {result['scrape']}")
    print(f"  NAME: {result['name']}")
    print(f"  Number of banks: {len(result['banks'])}")
    print("\nBanks found:")
    for i, bank in enumerate(result['banks'], 1):
        print(f"  {i}. {bank}")
    
    print("\n" + "=" * 70)
    print("Parser test completed!")

if __name__ == "__main__":
    test_list_parser()

