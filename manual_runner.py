"""
Manual runner for OJK Scraper
Runs the scraper once and saves output to specified directory
Use this for manual testing or one-time execution
"""

import sys
import logging
from pathlib import Path
from datetime import datetime

# Add the parent directory to path to import the module
sys.path.insert(0, str(Path(__file__).parent))

# Import using importlib to handle directory name with spaces
import importlib.util
module_path = Path(__file__).parent / "Laporan Publikasi BPR Konvensional" / "scraper.py"
spec = importlib.util.spec_from_file_location("scraper_module", module_path)
scraper_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(scraper_module)
OJKExtJSScraper = scraper_module.OJKExtJSScraper

# Configure logging
log_dir = Path(__file__).parent / "logs"
log_dir.mkdir(exist_ok=True)
log_file = log_dir / f"scheduled_run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger(__name__)


def main():
    """Main function for manual runs"""
    # Target output directory
    target_output_dir = Path(r"D:\APP\OSS\client\assets\publikasi")
    
    logger.info("=" * 60)
    logger.info("OJK Scraper - Manual Run")
    logger.info(f"Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info(f"Target output directory: {target_output_dir}")
    logger.info("=" * 60)
    
    # Create target directory if it doesn't exist
    try:
        target_output_dir.mkdir(parents=True, exist_ok=True)
        logger.info(f"[OK] Output directory ready: {target_output_dir}")
    except Exception as e:
        logger.error(f"[ERROR] Failed to create output directory: {e}")
        return 1
    
    scraper = None
    try:
        # Initialize scraper with headless mode for manual runs
        logger.info("[INFO] Initializing scraper...")
        scraper = OJKExtJSScraper(headless=True)
        
        # Override output directory
        scraper.output_dir = target_output_dir
        scraper.output_dir.mkdir(parents=True, exist_ok=True)
        logger.info(f"[OK] Output directory set to: {target_output_dir}")
        
        # Initialize and navigate
        scraper.initialize()
        scraper.navigate_to_page()
        
        # Run all phases
        logger.info("[INFO] Starting 3-phase scraping process...")
        scraper.run_all_phases()
        
        logger.info("=" * 60)
        logger.info("[OK] Manual run completed successfully!")
        logger.info(f"Completed at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        logger.info(f"Output saved to: {target_output_dir}")
        logger.info("=" * 60)
        
        return 0
        
    except Exception as e:
        logger.error(f"[ERROR] Manual run failed: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return 1
        
    finally:
        # Cleanup
        if scraper:
            try:
                logger.info("[INFO] Cleaning up resources...")
                scraper.cleanup(kill_processes=True)
                logger.info("[OK] Cleanup completed")
            except Exception as cleanup_error:
                logger.error(f"[ERROR] Cleanup error: {cleanup_error}")


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)

