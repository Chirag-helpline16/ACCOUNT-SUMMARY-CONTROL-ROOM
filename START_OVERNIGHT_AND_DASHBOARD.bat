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

start "Account Summary Overnight Worker" cmd /k python "%~dp0batch_account_summaries.py" --input "C:\Users\admin\Desktop\bank_trails" --database "%~dp0data\account_summaries.sqlite" --fast-reprocess-duplicates --audit-workers 16 --process-workers 4 --watch --keep-awake
timeout /t 3 /nobreak >nul
start "Account Summary Dashboard" cmd /k python "%~dp0account_summary_dashboard.py" --database "%~dp0data\account_summaries.sqlite" --host 127.0.0.1 --port 5002
timeout /t 3 /nobreak >nul
start "" "http://127.0.0.1:5002"

echo The overnight worker and local dashboard have been started.
echo Close this window when ready. Keep the two opened process windows running.
pause
