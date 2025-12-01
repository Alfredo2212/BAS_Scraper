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
    
    # Path to list file in queue folder
    queue_dir = Path(r"C:\Users\MSI\Desktop\OSS\client\assets-no_backup\sindikasi\queue")
    
    # Find list files matching pattern list_DD_MM_YYYY
    import re
    list_files = []
    if queue_dir.exists():
        for file in queue_dir.glob("list_*"):
            if re.match(r'list_\d{2}_\d{2}_\d{4}$', file.stem):
                list_files.append(file)
    
    if not list_files:
        print(f"Error: No list files found in queue folder: {queue_dir}")
        print("Expected format: list_DD_MM_YYYY (e.g., list_28_11_2025)")
        return
    
    # Use the most recent list file (or first one if multiple)
    list_file = sorted(list_files, key=lambda p: p.stat().st_mtime, reverse=True)[0]
    print(f"Found {len(list_files)} list file(s), using: {list_file.name}")
    
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
