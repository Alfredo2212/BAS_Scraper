# PowerShell script to setup Windows Task Scheduler for OJK Scraper
# Runs every Tuesday and Thursday at 15:00

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "OJK Scraper - Task Scheduler Setup" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Get the current directory (where this script is located)
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PythonScript = Join-Path $ScriptDir "scheduled_runner.py"

# Get Python executable
$PythonExe = (Get-Command python -ErrorAction SilentlyContinue).Source
if (-not $PythonExe) {
    Write-Host "[ERROR] Python not found in PATH" -ForegroundColor Red
    Write-Host "Please ensure Python is installed and added to PATH" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "[INFO] Python executable: $PythonExe" -ForegroundColor Green
Write-Host "[INFO] Script path: $PythonScript" -ForegroundColor Green
Write-Host ""

# Task names
$TaskNameTuesday = "OJK_Publikasi_Scraper_Tuesday"
$TaskNameThursday = "OJK_Publikasi_Scraper_Thursday"

# Check if tasks already exist and delete them
$TuesdayTask = Get-ScheduledTask -TaskName $TaskNameTuesday -ErrorAction SilentlyContinue
if ($TuesdayTask) {
    Write-Host "[INFO] Tuesday task already exists. Deleting..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $TaskNameTuesday -Confirm:$false
    Write-Host "[OK] Tuesday task deleted" -ForegroundColor Green
}

$ThursdayTask = Get-ScheduledTask -TaskName $TaskNameThursday -ErrorAction SilentlyContinue
if ($ThursdayTask) {
    Write-Host "[INFO] Thursday task already exists. Deleting..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $TaskNameThursday -Confirm:$false
    Write-Host "[OK] Thursday task deleted" -ForegroundColor Green
}

Write-Host ""

# Create action
$Action = New-ScheduledTaskAction -Execute $PythonExe -Argument "`"$PythonScript`"" -WorkingDirectory $ScriptDir

# Create trigger for Tuesday at 15:00
$TuesdayTrigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Tuesday -At "15:00"

# Create trigger for Thursday at 15:00
$ThursdayTrigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Thursday -At "15:00"

# Create settings
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable

# Create principal (run as SYSTEM)
$Principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

# Register Tuesday task
try {
    Register-ScheduledTask -TaskName $TaskNameTuesday -Action $Action -Trigger $TuesdayTrigger -Settings $Settings -Principal $Principal -Description "OJK Publikasi Scraper - Runs every Tuesday at 15:00" | Out-Null
    Write-Host "[OK] Tuesday task created successfully" -ForegroundColor Green
} catch {
    Write-Host "[ERROR] Failed to create Tuesday task: $_" -ForegroundColor Red
}

# Register Thursday task
try {
    Register-ScheduledTask -TaskName $TaskNameThursday -Action $Action -Trigger $ThursdayTrigger -Settings $Settings -Principal $Principal -Description "OJK Publikasi Scraper - Runs every Thursday at 15:00" | Out-Null
    Write-Host "[OK] Thursday task created successfully" -ForegroundColor Green
} catch {
    Write-Host "[ERROR] Failed to create Thursday task: $_" -ForegroundColor Red
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Setup Complete!" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Tasks created:" -ForegroundColor Green
Write-Host "  - $TaskNameTuesday  (Every Tuesday at 15:00)" -ForegroundColor Yellow
Write-Host "  - $TaskNameThursday (Every Thursday at 15:00)" -ForegroundColor Yellow
Write-Host ""
Write-Host "To view tasks:" -ForegroundColor Cyan
Write-Host "  Get-ScheduledTask -TaskName '$TaskNameTuesday'" -ForegroundColor White
Write-Host "  Get-ScheduledTask -TaskName '$TaskNameThursday'" -ForegroundColor White
Write-Host ""
Write-Host "To delete tasks:" -ForegroundColor Cyan
Write-Host "  Unregister-ScheduledTask -TaskName '$TaskNameTuesday' -Confirm:`$false" -ForegroundColor White
Write-Host "  Unregister-ScheduledTask -TaskName '$TaskNameThursday' -Confirm:`$false" -ForegroundColor White
Write-Host ""
Read-Host "Press Enter to exit"

