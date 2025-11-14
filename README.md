# OJK Publication Report Scraper

Automated web scraper for extracting financial publication reports from OJK (Otoritas Jasa Keuangan) website.

## Project Structure

```
Scraper/
├── config/                      # Configuration module
│   ├── __init__.py
│   └── settings.py             # Settings, constants, and selectors
├── scraper/                     # Main scraper module
│   ├── __init__.py
│   ├── ojk_scraper.py          # Main scraper orchestrator class
│   ├── selenium_setup.py       # WebDriver setup and configuration
│   ├── postback_handler.py     # ASP.NET PostBack event handling
│   ├── data_extractor.py       # HTML table data extraction
│   ├── excel_exporter.py       # Excel file export functionality
│   └── utils.py                # Utility functions
├── tests/                       # Step-by-step tests
│   ├── __init__.py
│   └── test_step1_setup.py     # Step 1: Setup verification
├── output/                      # Excel output files (gitignored)
├── logs/                        # Log files (gitignored)
├── main.py                      # Entry point for testing
├── requirements.txt             # Python dependencies
├── projectplan.md              # Detailed project plan
└── README.md
```

## Features

- **Modular Architecture**: Clean separation of concerns with dedicated modules
- **ASP.NET PostBack Handling**: Properly handles dynamic page reloads
- **Dynamic Element IDs**: Uses XPath with `contains()` for robust element finding
- **Rate Limiting**: Built-in delays and retry logic
- **Excel Export**: Automatic Excel file generation with formatting
- **Step-by-Step Testing**: Debug each component individually

## Setup

1. **Activate virtual environment** (Windows):
```bash
.\scraper_dev\Scripts\activate
```

2. **Install dependencies**:
```bash
pip install -r requirements.txt
```

3. **Verify setup**:
```bash
python tests/test_step1_setup.py
```

## Usage

### Test Step 1 (Setup Verification):
```bash
python tests/test_step1_setup.py
```

### Run Main Entry Point:
```bash
python main.py
```

## Implementation Steps

- [x] **Step 1**: Project structure and dependencies ✓
- [ ] **Step 2**: Selenium WebDriver setup and page navigation
- [ ] **Step 3**: Tab selection (BPR Konvensional)
- [ ] **Step 4**: Month dropdown with PostBack handling
- [ ] **Step 5**: Year dropdown with PostBack handling
- [ ] **Step 6**: Province dropdown with PostBack handling
- [ ] **Step 7**: City dropdown with PostBack handling
- [ ] **Step 8**: Bank dropdown with PostBack handling
- [ ] **Step 9**: Checkbox selection
- [ ] **Step 10**: Form submission and table waiting
- [ ] **Step 11**: Data extraction from tables
- [ ] **Step 12**: Excel export functionality

## Configuration

All settings are centralized in `config/settings.py`:
- Timeouts and delays
- XPath selectors
- User agents
- Output settings

## Notes

- The scraper handles ASP.NET PostBack events automatically
- Rate limiting is built-in (2-3 second delays between requests)
- Excel files are saved to the `output/` directory
- All selectors use XPath with `contains()` to handle dynamic IDs

