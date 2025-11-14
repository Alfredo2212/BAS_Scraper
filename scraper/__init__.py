"""
OJK Publication Report Scraper
Main scraper module for extracting financial reports from OJK website
"""

from scraper.ojk_scraper import OJKScraper
from scraper.selenium_setup import SeleniumSetup
from scraper.postback_handler import PostBackHandler
from scraper.data_extractor import DataExtractor
from scraper.excel_exporter import ExcelExporter

__version__ = "1.0.0"

__all__ = [
    'OJKScraper',
    'SeleniumSetup',
    'PostBackHandler',
    'DataExtractor',
    'ExcelExporter',
]

