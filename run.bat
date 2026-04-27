@echo off
echo Checking Python...
python --version 2>NUL
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo Download it from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during install!
    pause
    exit /b 1
)

echo Installing required packages...
python -m pip install playwright openpyxl
if errorlevel 1 (
    echo.
    echo ERROR: pip install failed. Make sure you have internet access.
    pause
    exit /b 1
)

echo.
echo Starting Excise Portal Scraper...
echo.
python excise_portal_scraper_v5.py
if errorlevel 1 (
    echo.
    echo ========================================
    echo  The app crashed. See error above.
    echo ========================================
    pause
)
