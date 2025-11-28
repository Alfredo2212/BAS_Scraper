"""
Shared scraper execution logic
Used by both manual_runner.py and scheduler_service.py
Ensures identical execution behavior
"""

import logging
from pathlib import Path
from datetime import datetime
from typing import Optional

logger = logging.getLogger(__name__)


def run_scraper_execution(
    scraper_class,
    target_output_dir: Path,
    run_type: str = "Run"
) -> bool:
    """
    Shared function to execute the scraper
    Used by both manual runner and scheduler service
    
    Args:
        scraper_class: The OJKExtJSScraper class (not instance)
        target_output_dir: Path to output directory
        run_type: Type of run for logging (e.g., "Manual Run", "Scheduled Job")
    
    Returns:
        bool: True if successful, False otherwise
    """
    scraper = None
    
    try:
        logger.info("=" * 70)
        logger.info(f"OJK Scraper - {run_type} Started")
        logger.info(f"Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        logger.info(f"Target output directory: {target_output_dir}")
        logger.info("=" * 70)
        
        # Create target directory if it doesn't exist
        try:
            target_output_dir.mkdir(parents=True, exist_ok=True)
            logger.info(f"[OK] Output directory ready: {target_output_dir}")
        except Exception as e:
            logger.error(f"[ERROR] Failed to create output directory: {e}")
            return False
        
        # Initialize scraper with headless mode
        logger.info("[INFO] Initializing scraper...")
        scraper = scraper_class(headless=True)
        
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
        
        logger.info("=" * 70)
        logger.info(f"[OK] {run_type} completed successfully!")
        logger.info(f"Completed at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        logger.info(f"Output saved to: {target_output_dir}")
        logger.info("=" * 70)
        
        return True
        
    except Exception as e:
        logger.error(f"[ERROR] {run_type} failed: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return False
        
    finally:
        # Cleanup
        if scraper:
            try:
                logger.info("[INFO] Cleaning up resources...")
                scraper.cleanup(kill_processes=True)
                logger.info("[OK] Cleanup completed")
            except Exception as cleanup_error:
                logger.error(f"[ERROR] Cleanup error: {cleanup_error}")

