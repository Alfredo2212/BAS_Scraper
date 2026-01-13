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

# Import shared execution function
from scraper_runner import run_scraper_execution

# Configure logging
log_dir = Path(__file__).parent / "logs"
log_dir.mkdir(exist_ok=True)
log_file = log_dir / f"manual_run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

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
    # Target output directory from environment variable
    from config.settings import Settings
    target_output_dir = Settings.OUTPUT_PUBLIKASI
    
    # Use shared execution function - ensures identical behavior with scheduler
    success = run_scraper_execution(
        scraper_class=OJKExtJSScraper,
        target_output_dir=target_output_dir,
        run_type="Manual Run"
    )
    
    return 0 if success else 1


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)

