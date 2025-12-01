"""
Test script for BPR Konvensional Retry - Targeted Scrape
Tests retry functionality for banks with zero values in existing Excel file
"""

import sys
import importlib.util
from pathlib import Path

# Import using importlib to handle directory name with spaces
module_path = Path(__file__).parent / "Laporan Publikasi BPR Konvensional" / "scraper.py"
spec = importlib.util.spec_from_file_location("publikasi_scraper", module_path)
publikasi_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(publikasi_module)
OJKExtJSScraper = publikasi_module.OJKExtJSScraper


def main():
    """Test the retry functionality"""
    # Check for headless flag from command line
    # Default to headless=True (no Chrome window) unless --visible or -v is passed
    headless = '--visible' not in sys.argv and '-v' not in sys.argv
    
    # Find Excel file in output/publikasi
    publikasi_dir = Path("output/publikasi")
    if not publikasi_dir.exists():
        print(f"Error: Directory not found: {publikasi_dir}")
        return
    
    # Find the most recent Excel file
    excel_files = list(publikasi_dir.glob("Publikasi_*.xlsx"))
    if not excel_files:
        print(f"Error: No Excel files found in {publikasi_dir}")
        return
    
    # Get the most recent file
    excel_file = max(excel_files, key=lambda p: p.stat().st_mtime)
    print(f"Found Excel file: {excel_file.name}")
    
    # Extract month and year from filename (format: Publikasi_MM_YYYY.xlsx)
    filename_parts = excel_file.stem.split('_')
    if len(filename_parts) < 3:
        print(f"Error: Cannot parse filename format: {excel_file.name}")
        print("Expected format: Publikasi_MM_YYYY.xlsx")
        return
    
    month_num = filename_parts[1]  # e.g., "09"
    year = filename_parts[2]  # e.g., "2025"
    
    # Convert month number to month name
    month_map = {
        "01": "Januari", "02": "Februari", "03": "Maret", "04": "April",
        "05": "Mei", "06": "Juni", "07": "Juli", "08": "Agustus",
        "09": "September", "10": "Oktober", "11": "November", "12": "Desember"
    }
    
    month_name = month_map.get(month_num, "September")
    
    print("=" * 70)
    print("BPR Konvensional Retry - Targeted Scrape Test")
    print("=" * 70)
    print(f"Excel file: {excel_file.name}")
    print(f"Month: {month_name} {year}")
    print(f"Headless mode: {headless}")
    print()
    
    # Initialize scraper (headless mode based on command line argument)
    scraper = OJKExtJSScraper(headless=headless)
    
    try:
        # Initialize browser
        scraper.initialize()
        
        # Run retry for zero value banks
        scraper._retry_zero_value_banks(month_name, year)
        
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

