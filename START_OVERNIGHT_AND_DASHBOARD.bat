@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>&1
if errorlevel 1 (
    echo Python was not found. Install Python or add it to PATH, then try again.
    pause
    exit /b 1
)

if not exist "%~dp0data" mkdir "%~dp0data"

set "WORKER_SCRIPT=%~dp0batch_account_summaries.py"
set "DASHBOARD_SCRIPT=%~dp0account_summary_dashboard.py"
set "DATABASE_PATH=%~dp0data\account_summaries.sqlite"
set "INPUT_PATH=C:\Users\admin\Desktop\bank_trails"

rem A normal start resumes pending SQLite rows.  Only the separate
rem REPROCESS_ALL_STRICT_ONCE.bat intentionally resets all files to pending.
powershell -NoProfile -Command "$target = [Regex]::Escape($env:WORKER_SCRIPT); if (Get-CimInstance Win32_Process | Where-Object { $_.Name -ieq 'python.exe' -and $_.CommandLine -match $target }) { exit 0 } else { exit 1 }"
if errorlevel 1 (
    echo Starting the overnight worker from its saved SQLite progress...
    start "Account Summary Overnight Worker" cmd /k python "%WORKER_SCRIPT%" --input "%INPUT_PATH%" --database "%DATABASE_PATH%" --process-workers 4 --watch --keep-awake
    timeout /t 3 /nobreak >nul
) else (
    echo The overnight worker is already running. Existing progress is continuing.
)

powershell -NoProfile -Command "if (Get-NetTCPConnection -LocalPort 5002 -State Listen -ErrorAction SilentlyContinue) { exit 0 } else { exit 1 }"
if errorlevel 1 (
    echo Starting the local dashboard...
    start "Account Summary Dashboard" cmd /k python "%DASHBOARD_SCRIPT%" --database "%DATABASE_PATH%" --host 127.0.0.1 --port 5002
    timeout /t 3 /nobreak >nul
) else (
    echo The dashboard is already running on port 5002.
)
start "" "http://127.0.0.1:5002"

echo The overnight worker and local dashboard are ready.
echo Completed files are never reset by this starter; pending files resume automatically.
pause
