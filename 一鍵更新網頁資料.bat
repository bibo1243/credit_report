@echo off
setlocal
set "REPO_DIR=C:\Users\feeling\Documents\GitHub\credit_report"
cd /d "%REPO_DIR%"

echo Updating donation directory data from Excel...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%REPO_DIR%\Update-DonationsData.ps1"
set "UPDATE_EXIT=%ERRORLEVEL%"

echo.
if not "%UPDATE_EXIT%"=="0" (
    echo Update failed. Please check the Excel shortcut or source file and try again.
    pause
    exit /b %UPDATE_EXIT%
)

echo Done. Reopen or refresh index.html.
pause
