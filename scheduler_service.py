"""
OJK Publikasi Scraper - Continuous Scheduler Service
Long-running service that runs the scraper automatically on schedule
Uses APScheduler to run every Tuesday and Thursday at 15:00 (GMT+7)
"""

import sys
import logging
from pathlib import Path
from datetime import datetime
from apscheduler.schedulers.blocking import BlockingScheduler
from apscheduler.triggers.cron import CronTrigger
import pytz

# Add the parent directory to path to import the module
sys.path.insert(0, str(Path(__file__).parent))

# Import using importlib to handle directory name with spaces
import importlib.util
module_path = Path(__file__).parent / "Laporan Publikasi BPR Konvensional" / "scraper.py"
spec = importlib.util.spec_from_file_location("scraper_module", module_path)
scraper_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(scraper_module)
OJKExtJSScraper = scraper_module.OJKExtJSScraper

# Import shared execution function
from scraper_runner import run_scraper_execution

### LOGGING SETUP ###
log_dir = Path(__file__).parent / "logs"
log_dir.mkdir(exist_ok=True)
log_file = log_dir / "scheduler.log"

# Configure logging to append to single log file
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8', mode='a'),
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger(__name__)

### JOB FUNCTION ###
def run_scraper_job():
    """
    Job function that runs the scraper
    Called by APScheduler at scheduled times
    Uses shared execution function to ensure identical behavior with manual runner
    """
    target_output_dir = Path(r"D:\APP\OSS\client\assets\publikasi")
    
    # Use shared execution function - ensures identical behavior with manual runner
    run_scraper_execution(
        scraper_class=OJKExtJSScraper,
        target_output_dir=target_output_dir,
        run_type="Scheduled Job"
    )

### SCHEDULER SETUP ###
def get_next_run_times(scheduler):
    """Calculate and return next scheduled run times"""
    jobs = scheduler.get_jobs()
    next_runs = []
    tz = pytz.timezone('Asia/Jakarta')
    now = datetime.now(tz)
    
    for job in jobs:
        try:
            # Use trigger's get_next_fire_time to calculate next run
            # First parameter is previous fire time (None for first run), second is now
            next_run = job.trigger.get_next_fire_time(None, now)
            if next_run:
                # Convert to Jakarta timezone if needed
                if next_run.tzinfo is None:
                    next_run = tz.localize(next_run)
                next_runs.append(next_run.strftime('%Y-%m-%d %H:%M:%S'))
        except Exception as e:
            logger.debug(f"Could not get next run time for job {job.id}: {e}")
    
    return sorted(next_runs) if next_runs else []

def main():
    """Main function - starts the continuous scheduler service"""
    logger.info("=" * 70)
    logger.info("OJK Publikasi Scraper - Scheduler Service Starting")
    logger.info(f"Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("=" * 70)
    
    # Create scheduler with GMT+7 timezone (Asia/Jakarta)
    scheduler = BlockingScheduler(timezone=pytz.timezone('Asia/Jakarta'))
    
    # Schedule job for Tuesday at 15:00
    scheduler.add_job(
        run_scraper_job,
        trigger=CronTrigger(day_of_week='tue', hour=15, minute=0),
        id='ojk_scraper_tuesday',
        name='OJK Scraper - Tuesday 15:00',
        replace_existing=True
    )
    
    # Schedule job for Thursday at 15:00
    scheduler.add_job(
        run_scraper_job,
        trigger=CronTrigger(day_of_week='thu', hour=15, minute=0),
        id='ojk_scraper_thursday',
        name='OJK Scraper - Thursday 15:00',
        replace_existing=True
    )
    
    # Get and display next run times
    try:
        next_runs = get_next_run_times(scheduler)
        logger.info("Scheduler Started - waiting for Tue/Thu 15:00")
        if next_runs:
            logger.info("Next scheduled runs:")
            for next_run in next_runs:
                logger.info(f"  - {next_run}")
        else:
            logger.info("Next scheduled runs: (calculating...)")
    except Exception as e:
        logger.warning(f"Could not calculate next run times: {e}")
        logger.info("Scheduler Started - waiting for Tue/Thu 15:00")
    
    logger.info("=" * 70)
    logger.info("Scheduler is running. Press Ctrl+C to stop.")
    logger.info("=" * 70)
    
    try:
        # Start the scheduler (this will block and run forever)
        scheduler.start()
    except (KeyboardInterrupt, SystemExit):
        logger.info("=" * 70)
        logger.info("Scheduler service stopped by user")
        logger.info(f"Stopped at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        logger.info("=" * 70)
        scheduler.shutdown()
    except Exception as e:
        logger.error(f"[ERROR] Scheduler service error: {e}")
        import traceback
        logger.error(traceback.format_exc())
        scheduler.shutdown()
        sys.exit(1)

if __name__ == "__main__":
    main()

