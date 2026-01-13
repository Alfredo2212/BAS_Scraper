"""
Test script for IBPRS full scraping
Tests scraping all pages and saving to .txt and .xlsx files
"""

import sys
import importlib.util
from pathlib import Path

# Add the project root to path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

# Import the scraper module (handle spaces in folder name)
scraper_path = project_root / "Laporan Bulanan IBPRS" / "scraper.py"
spec = importlib.util.spec_from_file_location("ibprs_scraper", scraper_path)
ibprs_scraper_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(ibprs_scraper_module)
IBPRSScraper = ibprs_scraper_module.IBPRSScraper


def main():
    """Test IBPRS full scraping workflow"""
    print("=" * 70)
    print("IBPRS Full Scraping Test")
    print("=" * 70)
    print("This will:")
    print("1. Navigate to https://ibpr-s.ojk.go.id/DataKeuangan")
    print("2. Input 'Kep. Riau' in the input field")
    print("3. Click the 'Cari' button")
    print("4. Scrape all pages using BeautifulSoup")
    print("5. Click 'next' button until no more pages")
    print("6. Save results to output/IBPRS/ (both .txt and .xlsx)")
    print("=" * 70)
    print()
    
    # Create scraper instance (headless mode)
    scraper = IBPRSScraper(headless=True)
    
    try:
        # Run the complete scrape and save workflow
        success = scraper.scrape_and_save("Kep. Riau")
        
        if success:
            print()
            print("=" * 70)
            print("Scraping completed successfully!")
            print("Check output/IBPRS/ directory for .txt and .xlsx files")
            print("=" * 70)
        else:
            print()
            print("=" * 70)
            print("Scraping failed. Check logs for details.")
            print("=" * 70)
        
    except KeyboardInterrupt:
        print("\n\nScraping interrupted by user")
    except Exception as e:
        print(f"\n\nError during scraping: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # Cleanup is handled in scrape_and_save method
        print("Test complete.")


if __name__ == "__main__":
    main()

