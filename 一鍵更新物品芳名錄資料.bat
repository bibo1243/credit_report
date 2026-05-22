@echo off
setlocal
cd /d "%~dp0"

echo Updating goods directory data from Excel...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Update-GoodsData.ps1"
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if not "%EXIT_CODE%"=="0" (
    echo Update failed. Check whether the Excel file is locked or the headers changed.
    pause
    exit /b %EXIT_CODE%
)

echo Done. Reopen or refresh goods.html.
pause
