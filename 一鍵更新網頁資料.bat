@echo off
setlocal
cd /d "%~dp0"

echo Updating website data from Excel...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Update-DonationsData.ps1"
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if not "%EXIT_CODE%"=="0" (
    echo Update failed. Check whether the Excel file is locked or the headers changed.
    pause
    exit /b %EXIT_CODE%
)

echo Done. Reopen or refresh index.html.
pause
