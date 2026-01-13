"""
Test script for IBPRS navigation
Tests navigation to DataKeuangan page, input province, and click search button
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
    """Test IBPRS navigation flow"""
    print("=" * 70)
    print("IBPRS Navigation Test")
    print("=" * 70)
    print("This will:")
    print("1. Navigate to https://ibpr-s.ojk.go.id/DataKeuangan")
    print("2. Input 'Kep. Riau' in the input field")
    print("3. Click the 'Cari' button")
    print("=" * 70)
    print()
    
    # Create scraper instance (non-headless so we can see what's happening)
    scraper = IBPRSScraper(headless=False)
    
    try:
        # Run the navigation test
        success = scraper.test_navigation("Kep. Riau")
        
        if success:
            print()
            print("=" * 70)
            print("Navigation test completed successfully!")
            print("Browser will remain open for 5 seconds to view results.")
            print("=" * 70)
            # The test_navigation method already waits 5 seconds, no need for extra wait
        else:
            print()
            print("=" * 70)
            print("Navigation test failed. Check logs for details.")
            print("=" * 70)
        
    except KeyboardInterrupt:
        print("\n\nTest interrupted by user")
    except Exception as e:
        print(f"\n\nError during test: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # Cleanup
        scraper.cleanup()
        print("Browser closed. Test complete.")


if __name__ == "__main__":
    main()

