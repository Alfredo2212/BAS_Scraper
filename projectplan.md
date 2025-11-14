# Implement OJK Publication Report Scraping with ASP.NET PostBack Handling

## Overview

Implement automated web scraping for OJK (Otoritas Jasa Keuangan) publication reports. The system must handle ASP.NET PostBack events, dynamic element IDs, conditional table rendering, and rate limiting. The system will navigate to the OJK website, interact with form elements (dropdowns, checkboxes), submit the form, and extract the resulting data from scrollable table containers.

## Feasibility Assessment

**100% Feasible** - Fully automatable with Selenium WebDriver, but requires careful handling of:

- ASP.NET PostBack events
- Dynamic element IDs
- Conditional table rendering
- Rate limiting (10-20 requests)
- Scrollable table containers
- Dynamic CSS classes

## Critical Technical Challenges & Solutions

### 1. ASP.NET PostBack Handling (CRITICAL)

**Challenge**: OJK uses ASP.NET PostBack - dropdown selections trigger `__doPostBack()` JavaScript, causing partial page reloads.

**Solution**:

- Use `WebDriverWait` + `expected_conditions` to wait for:
  - New dropdowns to appear after postback
  - New select options to load
  - Page state to stabilize
- Wait for element staleness, then wait for new element presence
- Add 2-3 second delays between dropdown selections

**Implementation**:

```python
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

# Wait for postback completion
wait = WebDriverWait(driver, 30)
wait.until(EC.staleness_of(old_element))
wait.until(EC.presence_of_element_located((By.XPATH, "//select[contains(@id,'ddlKota')]")))
```

### 2. Dynamic Element IDs (CRITICAL)

**Challenge**: Element IDs change at runtime (e.g., `ctl00$ctl00$ContentPlaceHolder1$ContentPlaceHolderDetail$ddlProvinsi`)

**Solution**: DO NOT use ID selectors. Use:

- XPath with `contains()`: `//select[contains(@id,'ddlProvinsi')]`
- CSS based on structure
- Text-based selectors

**Implementation**:

```python
# ❌ BAD - Don't use direct ID
provinsi_dd = driver.find_element(By.ID, "ctl00$ctl00$ContentPlaceHolder1$ContentPlaceHolderDetail$ddlProvinsi")

# ✅ GOOD - Use contains() XPath
provinsi_dd = driver.find_element(By.XPATH, "//select[contains(@id,'ddlProvinsi')]")
```

### 3. Conditional Table Rendering

**Challenge**: Table only appears AFTER clicking "Tampilkan" with correct combination. If bank has no data → No table is shown.

**Solution**: Use try/except to detect table presence, skip gracefully if not found.

**Implementation**:

```python
try:
    table_container = wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@id,'table')]")))
    # Extract data
except TimeoutException:
    # No data available for this combination
    return {"success": False, "message": "No data available"}
```

### 4. Table Structure & Parsing

**Challenge**:

- Table renders inside iframe-like DIV (not JSON)
- Results are HTML table with inline-styled `<div>` and `<td>`
- Table may be in scrollable container
- Dynamic CSS classes (e.g., `A1a...` random hash) - cannot rely on class names

**Solution**:

- Use BeautifulSoup to parse `page_source` after obtaining HTML
- Get `innerHTML` of scroll container, not whole page
- Rely on `<tr>` count and structure instead of CSS classes

**Implementation**:

```python
from bs4 import BeautifulSoup

# Get innerHTML of scrollable container
table_html = table_container.get_attribute('innerHTML')
soup = BeautifulSoup(table_html, 'html.parser')

# Parse using <tr> structure (not CSS classes)
rows = soup.find_all('tr')
for row in rows:
    cells = row.find_all('td')
    # Extract data based on position
```

### 5. Rate Limiting & Session Management

**Challenge**: OJK rate-limits after 10-20 requests. May get: blank results, session reset, page refresh, element not found.

**Solution**:

- Add 2-3 second delays between requests
- Rotate user-agents
- Implement retry logic with exponential backoff
- Handle session expiration gracefully

**Implementation**:

```python
import time
import random

# Add delay between requests
time.sleep(2 + random.uniform(0, 1))

# Rotate user agents
user_agents = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    # ... more user agents
]
options.add_argument(f'user-agent={random.choice(user_agents)}')
```

## Implementation Details

### 1. Add Dependencies

**File**: `requirements.txt`

- Add `selenium==4.15.0` (or latest stable version)
- Add `beautifulsoup4==4.12.0` (for HTML parsing)
- Add `webdriver-manager==4.0.0` (for automatic WebDriver management)

### 2. Backend Router Implementation

**File**: `backend/app/routers/publikasi.py`

**Endpoints to create:**

- `POST /publikasi/scrape` - Main scraping endpoint
  - Parameters: period (month, year), province, city, bank, report types
  - Returns: Scraped data in structured format (JSON)

**Implementation approach:**

1. Initialize Selenium WebDriver with Chrome headless mode
2. Navigate to: `https://ojk.go.id/id/kanal/perbankan/data-dan-statistik/laporan-keuangan-perbankan/Default.aspx`
3. Wait for page to fully load
4. Select "BPR Konvensional" tab (wait for tab content to load)
5. Fill form fields with postback handling:

   - Select month dropdown → Wait for postback completion
   - Select year dropdown → Wait for postback completion
   - Select province dropdown → Wait for postback + city dropdown to populate
   - Select city/regency dropdown → Wait for postback + bank dropdown to populate
   - Select bank dropdown → Wait for postback completion
   - Tick first 3 checkboxes (Laporan Posisi Keuangan, Laporan Laba Rugi, Laporan Kualitas Aset Produktif)

6. Click "Tampilkan" button
7. Wait for results table container to appear (with timeout)
8. Extract innerHTML from scrollable table container
9. Parse HTML table with BeautifulSoup using `<tr>` structure
10. Convert to structured data (list of dicts)
11. Return structured data or handle "no data" case gracefully

**Error handling:**

- Network timeouts
- PostBack timeout (element not appearing after selection)
- Element not found (use robust selectors)
- Rate limiting (detect blank results, implement backoff)
- No data available (table not rendered - skip gracefully)
- Session expiration (detect and re-authenticate if needed)

### 3. UI Window Implementation

**File**: `client/windows/publikasi_ui.py`

**Features:**

- Form window for user input:
  - Period selection (month, year dropdowns)
  - Province dropdown (populated from OJK or predefined list)
  - City/Regency dropdown (dynamic based on province)
  - Bank dropdown (dynamic based on city)
  - Report type checkboxes (first 3 pre-selected)
- "Simpan Laporan Publikasi" button
- Progress indicator during scraping
- Display results in a table/treeview
- Export functionality (Excel/CSV)

### 4. API Client Functions

**File**: `client/api/publikasi.py` (new file)

**Functions:**

- `scrape_publication_report(params)` - Call backend scraping endpoint
- `get_provinces()` - Get available provinces (if needed)
- `get_cities(province)` - Get cities for a province
- `get_banks(city)` - Get banks for a city

### 5. Update Main Window

**File**: `client/windows/me_window.py`

**Update `capture_publication_report()` function:**

- Replace placeholder with actual implementation
- Open publikasi_ui window for user input
- Handle the scraping process
- Display results

### 6. Database Schema (Optional - for storing scraped data)

**Considerations:**

- Table: `hr.publikasi_reports` or similar
- Store: period, bank, report_type, scraped_data, scraped_date
- This allows comparison between different periods

## Files to Create/Modify

1. **New Files:**

   - `backend/app/routers/publikasi.py` - Backend router
   - `client/windows/publikasi_ui.py` - UI window
   - `client/api/publikasi.py` - API client functions

2. **Modified Files:**

   - `requirements.txt` - Add Selenium and dependencies
   - `client/windows/me_window.py` - Update capture_publication_report()
   - `backend/app/main.py` - Register publikasi router (if needed)

## Testing Strategy

1. Test with different periods (months/years)
2. Test with different provinces/cities
3. Test with different banks
4. Test error scenarios:

   - Network failure
   - PostBack timeout
   - Element not found
   - No data available
   - Rate limiting

5. Test data extraction accuracy
6. Test with scrollable tables (large datasets)

## Notes

- Ensure compliance with OJK website terms of service
- Implement rate limiting to avoid overwhelming the server (2-3 second delays)
- Handle cookies/sessions for stateful interactions
- Implement retry logic for transient failures
- Use headless browser mode for server-side execution
- Implement timeout limits (e.g., 90 seconds max for full flow)
- Consider caching province/city/bank lists (if stable)
- Use async/background processing for long-running operations