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

echo Publishing donation directory website...
echo.

if exist ".git\index.lock" del /f /q ".git\index.lock"

git add index.html donations.json donations-data.js logo.png Update-DonationsData.ps1
if errorlevel 1 (
    echo Failed to stage donation website files.
    pause
    exit /b 1
)

git diff --cached --quiet
if not errorlevel 1 goto push_only

git commit -m "Publish donation directory update"
if errorlevel 1 (
    echo Commit failed.
    pause
    exit /b 1
)

:push_only
git push origin main
if errorlevel 1 (
    echo.
    echo Publish failed. If a GitHub login window appears, please finish login and run this file again.
    pause
    exit /b 1
)

echo.
echo Update and publish complete.
echo Website URL:
echo https://bibo1243.github.io/credit_report/
pause
