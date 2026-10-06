@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo   Amazon Seller Lead Crawler
echo ========================================
echo.
if not exist "venv" (
    echo Installing dependencies...
    pip install -r requirements.txt
)
echo Starting crawler...
python amazon_crawler.py
pause
