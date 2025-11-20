# OJK Scraper Scheduler Setup

This document explains how to set up the automated scheduler for the OJK Publikasi Scraper.

## Overview

The scheduler runs the scraper automatically every **Tuesday and Thursday at 15:00** and saves the Excel files to:
```
D:\APP\OSS\client\assets\publikasi
```

## Files

- **`scheduled_runner.py`**: Main script that runs the scraper (called by Task Scheduler)
- **`setup_scheduler.bat`**: Batch script to set up Windows Task Scheduler (Command Prompt)
- **`setup_scheduler.ps1`**: PowerShell script to set up Windows Task Scheduler (PowerShell)

## Setup Instructions

### Option 1: Using PowerShell (Recommended)

1. Open PowerShell as **Administrator**
2. Navigate to the project directory:
   ```powershell
   cd "C:\Users\MSI\Desktop\Scraper"
   ```
3. Run the setup script:
   ```powershell
   .\setup_scheduler.ps1
   ```

### Option 2: Using Command Prompt

1. Open Command Prompt as **Administrator**
2. Navigate to the project directory:
   ```cmd
   cd "C:\Users\MSI\Desktop\Scraper"
   ```
3. Run the setup script:
   ```cmd
   setup_scheduler.bat
   ```

## What Gets Created

The setup script creates two Windows Scheduled Tasks:

1. **`OJK_Publikasi_Scraper_Tuesday`** - Runs every Tuesday at 15:00
2. **`OJK_Publikasi_Scraper_Thursday`** - Runs every Thursday at 15:00

Both tasks:
- Run as SYSTEM account (no user login required)
- Run in headless mode (no browser window)
- Save output to `D:\APP\OSS\client\assets\publikasi`
- Create log files in `logs\scheduled_run_YYYYMMDD_HHMMSS.log`

## Verify Setup

### Using PowerShell:
```powershell
Get-ScheduledTask -TaskName "OJK_Publikasi_Scraper_Tuesday"
Get-ScheduledTask -TaskName "OJK_Publikasi_Scraper_Thursday"
```

### Using Command Prompt:
```cmd
schtasks /query /tn "OJK_Publikasi_Scraper_Tuesday"
schtasks /query /tn "OJK_Publikasi_Scraper_Thursday"
```

### Using Task Scheduler GUI:
1. Open **Task Scheduler** (search in Start menu)
2. Look for tasks starting with `OJK_Publikasi_Scraper_`

## Manual Test Run

To test the scheduled runner manually:

```powershell
python scheduled_runner.py
```

Or:

```cmd
python scheduled_runner.py
```

## View Logs

Log files are created in the `logs` directory with the format:
```
logs/scheduled_run_YYYYMMDD_HHMMSS.log
```

## Remove Scheduled Tasks

### Using PowerShell:
```powershell
Unregister-ScheduledTask -TaskName "OJK_Publikasi_Scraper_Tuesday" -Confirm:$false
Unregister-ScheduledTask -TaskName "OJK_Publikasi_Scraper_Thursday" -Confirm:$false
```

### Using Command Prompt:
```cmd
schtasks /delete /tn "OJK_Publikasi_Scraper_Tuesday" /f
schtasks /delete /tn "OJK_Publikasi_Scraper_Thursday" /f
```

## Troubleshooting

### Task doesn't run
1. Check if Python is in PATH:
   ```cmd
   python --version
   ```
2. Check task history in Task Scheduler GUI
3. Check log files in `logs` directory
4. Verify the output directory exists: `D:\APP\OSS\client\assets\publikasi`

### Permission errors
- Make sure you run the setup script as Administrator
- Check that the output directory has write permissions

### Python not found
- Ensure Python is installed and added to system PATH
- You can specify full Python path in the setup scripts if needed

## Notes

- The scraper runs in **headless mode** (no visible browser window)
- Each run creates a new log file with timestamp
- Excel files are saved with format: `Publikasi_MM_YY.xlsx`
- The scraper uses the 3-phase approach to prevent "Bad Request" errors

