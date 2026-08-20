@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>&1
if errorlevel 1 (
    echo Python was not found. Install Python or add it to PATH, then try again.
    pause
    exit /b 1
)

echo This will recalculate every source workbook with credit-only duplicate handling.
echo Existing summaries remain available until each workbook is atomically replaced.
echo.
python "%~dp0batch_account_summaries.py" --input "C:\Users\admin\Desktop\bank_trails" --database "%~dp0data\account_summaries.sqlite" --reprocess-all --watch --keep-awake

echo.
echo The strict reprocessing worker stopped. Progress is saved.
pause
