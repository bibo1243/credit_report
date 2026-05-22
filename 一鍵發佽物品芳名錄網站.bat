@echo off
setlocal
cd /d "%~dp0"

echo Publishing goods directory website...
echo.

git add goods.html goods.json goods-data.js Update-GoodsData.ps1 "一鍵更新物品芳名錄資料.bat" "一鍵發佽物品芳名錄網站.bat"
if errorlevel 1 (
    echo Failed to stage goods website files.
    pause
    exit /b 1
)

git diff --cached --quiet
if not errorlevel 1 goto push_only

git commit -m "Publish goods directory update"
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
echo Publish complete.
echo Website URL:
echo https://bibo1243.github.io/credit_report/goods.html
pause
