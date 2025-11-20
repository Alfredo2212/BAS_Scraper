@echo off
REM Setup Windows Task Scheduler for OJK Scraper
REM Runs every Tuesday and Thursday at 15:00

echo ========================================
echo OJK Scraper - Task Scheduler Setup
echo ========================================
echo.

REM Get the current directory (where this script is located)
set SCRIPT_DIR=%~dp0
set PYTHON_SCRIPT=%SCRIPT_DIR%scheduled_runner.py

REM Get Python executable
where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python not found in PATH
    echo Please ensure Python is installed and added to PATH
    pause
    exit /b 1
)

for /f "delims=" %%i in ('where python') do set PYTHON_EXE=%%i

echo [INFO] Python executable: %PYTHON_EXE%
echo [INFO] Script path: %PYTHON_SCRIPT%
echo.

REM Task name
set TASK_NAME=OJK_Publikasi_Scraper

REM Check if task already exists
schtasks /query /tn "%TASK_NAME%" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [INFO] Task already exists. Deleting existing task...
    schtasks /delete /tn "%TASK_NAME%" /f
    echo [OK] Existing task deleted
    echo.
)

REM Create the scheduled task
echo [INFO] Creating scheduled task...
echo Task Name: %TASK_NAME%
echo Schedule: Every Tuesday and Thursday at 15:00
echo.

REM Create task for Tuesday at 15:00
schtasks /create /tn "%TASK_NAME%_Tuesday" /tr "\"%PYTHON_EXE%\" \"%PYTHON_SCRIPT%\"" /sc weekly /d TUE /st 15:00 /ru "SYSTEM" /f
if %ERRORLEVEL% EQU 0 (
    echo [OK] Tuesday task created successfully
) else (
    echo [ERROR] Failed to create Tuesday task
)

REM Create task for Thursday at 15:00
schtasks /create /tn "%TASK_NAME%_Thursday" /tr "\"%PYTHON_EXE%\" \"%PYTHON_SCRIPT%\"" /sc weekly /d THU /st 15:00 /ru "SYSTEM" /f
if %ERRORLEVEL% EQU 0 (
    echo [OK] Thursday task created successfully
) else (
    echo [ERROR] Failed to create Thursday task
)

echo.
echo ========================================
echo Setup Complete!
echo ========================================
echo.
echo Tasks created:
echo   - %TASK_NAME%_Tuesday  (Every Tuesday at 15:00)
echo   - %TASK_NAME%_Thursday (Every Thursday at 15:00)
echo.
echo To view tasks: schtasks /query /tn "%TASK_NAME%_Tuesday"
echo To delete tasks: schtasks /delete /tn "%TASK_NAME%_Tuesday" /f
echo                  schtasks /delete /tn "%TASK_NAME%_Thursday" /f
echo.
pause

