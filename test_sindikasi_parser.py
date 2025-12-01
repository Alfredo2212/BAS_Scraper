"""
Test script for Sindikasi Scraper - Parser
Tests parsing of Syariah form 1 data
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


def main():
    """Test the parser functionality"""
    # Check for headless flag from command line
    # Default to headless=True (no Chrome window) unless --visible or -v is passed
    headless = '--visible' not in sys.argv and '-v' not in sys.argv
    
    # Path to list file
    list_file = Path("input/list_28_11_2025")
    
    if not list_file.exists():
        print(f"Error: List file not found: {list_file}")
        return
    
    print("=" * 70)
    print("Sindikasi Scraper - Parser Test")
    print("=" * 70)
    print(f"List file: {list_file}")
    print(f"Headless mode: {headless}")
    print()
    
    # Initialize scraper (headless mode based on command line argument)
    scraper = SindikasiScraper(headless=headless)
    
    try:
        # Run the scraper
        scraper.find_all_banks(list_file)
        
        print()
        print("=" * 70)
        print("Test completed!")
        print("=" * 70)
        
    except KeyboardInterrupt:
        print("\n\nTest interrupted by user")
        scraper.cleanup()
    except Exception as e:
        print(f"\n\nError during test: {e}")
        import traceback
        traceback.print_exc()
        scraper.cleanup()
    finally:
        if scraper.driver:
            scraper.cleanup()


if __name__ == "__main__":
    main()
