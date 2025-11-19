# Upgrade Scraper to Support 3 Separate Phases

## Overview

Add minimal changes to existing `scrape_all_data()` to support running in 3 separate phases, each with its own Chrome session. Preserve all existing core logic - only add phase parameter and orchestration.

## Current Structure (Preserved)

- Checkbox 001: `treeview-1012-record-BPK-901-000001` → Sheets 1-3 (ASET, Kredit, DPK)
- Checkbox 002: `treeview-1012-record-BPK-901-000002` → Sheet 4 (Laba Kotor)  
- Checkbox 003: `treeview-1012-record-BPK-901-000003` → Sheet 5 (Rasio)

## Month Selection Logic (Update)

Based on current server date:

- Jan, Feb, Mar → use December (YYYY-1)
- Apr, May, Jun → use March YYYY
- Jul, Aug, Sep → use June YYYY
- Oct, Nov, Dec → use September YYYY
Where YYYY is the current server year.

## Implementation Plan

### 1. Update Month Selection Logic

**File: `Laporan Publikasi BPR Konvensional/scraper.py`**

#### 1.1 Update `_get_target_month_year()` method

- Add logic to map current month to quarterly report month (preserve existing logic, add quarterly mapping)
- Current month Jan/Feb/Mar → return December of previous year
- Current month Apr/May/Jun → return March of current year
- Current month Jul/Aug/Sep → return June of current year
- Current month Oct/Nov/Dec → return September of current year

### 2. Add Phase Parameter to Existing Methods (Minimal Changes)

#### 2.1 Add `phase` parameter to `scrape_all_data(month, year, phase='all')`

- Keep all existing logic intact
- `phase='all'`: Run all 3 phases sequentially (current behavior - preserve exactly as is)
- `phase='001'`: Run only Phase 1 (checkbox 001 → Sheets 1-3), then close Chrome and update Excel
- `phase='002'`: Run only Phase 2 (checkbox 002 → Sheet 4), then close Chrome and update Excel
- `phase='003'`: Run only Phase 3 (checkbox 003 → Sheet 5), then close Chrome and update Excel

#### 2.2 Modify checkbox selection based on phase

- If `phase='001'`: Use existing `_select_initial_dropdowns_and_checkboxes()` (checkbox 001)
- If `phase='002'`: Create new method `_select_checkbox_002_only()` (similar to existing `_change_checkboxes_for_laba_kotor()` but only check 002)
- If `phase='003'`: Create new method `_select_checkbox_003_only()` (similar to existing but only check 003)
- If `phase='all'`: Keep existing flow (checkbox 001, then refresh and checkboxes 002+003)

#### 2.3 Update iteration logic based on phase

- Phase 001: Use existing iteration with `extract_mode='sheets_1_3'` (preserve exactly)
- Phase 002: Use existing iteration with `extract_mode='sheets_4_5'` but skip Rasio extraction
- Phase 003: Use existing iteration with `extract_mode='sheets_4_5'` but skip Laba Kotor extraction

#### 2.4 Add Chrome close and Excel update after each phase

- After Phase 001 iteration completes: `self.cleanup()` then `_finalize_excel(month, year)`
- After Phase 002 iteration completes: `self.cleanup()` then `_finalize_excel_laba_kotor(month, year)`
- After Phase 003 iteration completes: `self.cleanup()` then `_finalize_excel_rasio(month, year)`

### 3. Create Main Orchestrator Method

#### 3.1 Create `run_all_phases(month=None, year=None)`

- Get month/year using updated `_get_target_month_year()` if not provided
- Call `scrape_all_data(month, year, phase='001')` → closes Chrome, updates Excel
- Call `scrape_all_data(month, year, phase='002')` → closes Chrome, updates Excel  
- Call `scrape_all_data(month, year, phase='003')` → closes Chrome, updates Excel
- Print completion message

### 4. Create Minimal Checkbox Selection Helpers

#### 4.1 Create `_select_checkbox_002_only()`

- Based on existing `_change_checkboxes_for_laba_kotor()` logic
- Only check checkbox 002, ensure 001 and 003 are unchecked

#### 4.2 Create `_select_checkbox_003_only()`

- Based on existing `_change_checkboxes_for_laba_kotor()` logic
- Only check checkbox 003, ensure 001 and 002 are unchecked

### 5. Preserve Existing Code

- Keep `_setup_for_sheets_4_5()` method (used when phase='all')
- Keep page refresh logic (used when phase='all')
- Keep `_change_checkboxes_for_laba_kotor()` (used when phase='all')
- Only add new methods, don't remove existing ones
- Excel finalization methods already work independently - no changes needed