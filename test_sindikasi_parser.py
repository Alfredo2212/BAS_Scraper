"""
Test script for Sindikasi Scraper - Parser
Tests parsing of Syariah form 1 data
"""

import sys
import importlib.util
import re
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
    
    # Find sindikasi files matching pattern sindikasi_NAME_DD_MM_YYYY.txt
    list_files = []
    if queue_dir.exists():
        for file in queue_dir.glob("sindikasi_*.txt"):
            if re.match(r'sindikasi_.+_\d{2}_\d{2}_\d{4}\.txt$', file.name):
                list_files.append(file)
    
    if not list_files:
        print(f"Error: No sindikasi queue files found in queue folder: {queue_dir}")
        print("Expected format: sindikasi_NAME_DD_MM_YYYY.txt (e.g., sindikasi_TestName_28_11_2025.txt)")
        return
    
    print(f"Found {len(list_files)} sindikasi file(s) in queue folder")
    
    # Process each file with SCRAPE = TRUE
    processed_count = 0
    for list_file in sorted(list_files):
        should_set_false = False
        scraper = None
        try:
            # Read file to check SCRAPE flag
            with open(list_file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Check if SCRAPE = TRUE
            scrape_match = re.search(r'SCRAPE\s*=\s*(TRUE|FALSE)', content, re.IGNORECASE)
            if not scrape_match or scrape_match.group(1).upper() != 'TRUE':
                print(f"[SKIP] File {list_file.name} has SCRAPE = FALSE, skipping")
                continue
            
            # Mark that we should set FALSE after processing
            should_set_false = True
            processed_count += 1
            
            print("=" * 70)
            print(f"Sindikasi Scraper - Parser Test ({processed_count}/{len(list_files)})")
            print("=" * 70)
            print(f"List file: {list_file}")
            print(f"Headless mode: {headless}")
            print()
            
            # Initialize scraper (headless mode based on command line argument)
            scraper = SindikasiScraper(headless=headless)
            
            # Clear any previous data and initialize
            scraper.all_bank_data = []  # Clear previous data
            scraper.initialize()  # Initialize browser
            
            # Run the scraper
            scraper.find_all_banks(list_file)
            
            print()
            print("=" * 70)
            print(f"Test completed for {list_file.name}!")
            print("=" * 70)
            
        except KeyboardInterrupt:
            print("\n\nTest interrupted by user")
            if scraper:
                scraper.cleanup()
            break
        except Exception as e:
            print(f"\n\nError during test for {list_file.name}: {e}")
            import traceback
            traceback.print_exc()
            if scraper:
                scraper.cleanup()
        finally:
            if scraper and scraper.driver:
                try:
                    scraper.cleanup()
                except:
                    pass
            
            # Set SCRAPE = FALSE after completion (success or failure)
            if should_set_false:
                try:
                    update_scrape_flag(list_file, False)
                    print(f"\n[OK] Set SCRAPE = FALSE in {list_file.name}")
                except Exception as e:
                    print(f"\n[WARNING] Failed to update SCRAPE flag in {list_file.name}: {e}")
    
    if processed_count == 0:
        print("\n[INFO] No files with SCRAPE = TRUE found to process")
    else:
        print(f"\n[INFO] Processed {processed_count} file(s) with SCRAPE = TRUE")


def update_scrape_flag(txt_file: Path, value: bool):
    """
    Update SCRAPE flag in the txt file
    
    Args:
        txt_file: Path to the txt file
        value: True or False
    """
    try:
        # Read file content
        with open(txt_file, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        # Update SCRAPE line
        updated_lines = []
        scrape_found = False
        for line in lines:
            if re.match(r'^\s*SCRAPE\s*=', line, re.IGNORECASE):
                updated_lines.append(f"SCRAPE = {'TRUE' if value else 'FALSE'}\n")
                scrape_found = True
            else:
                updated_lines.append(line)
        
        # If SCRAPE line not found, add it at the beginning
        if not scrape_found:
            updated_lines.insert(0, f"SCRAPE = {'TRUE' if value else 'FALSE'}\n")
        
        # Write back to file
        with open(txt_file, 'w', encoding='utf-8') as f:
            f.writelines(updated_lines)
        
    except Exception as e:
        print(f"Error updating SCRAPE flag: {e}")
        raise


if __name__ == "__main__":
    main()
