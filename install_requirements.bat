@echo off
chcp 65001 >nul
title Installing Patent Downloader Requirements
color 0B

echo ================================================
echo   Installing Required Packages
echo ================================================
echo.

REM Check Python version
python --version
if %errorlevel% neq 0 (
    echo ERROR: Python not found!
    echo Please install Python from: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo.
echo This will install the following packages:
echo   - pandas (for reading Excel files)
echo   - openpyxl (for Excel support)
echo   - selenium (for browser automation)
echo   - requests (for HTTP downloads)
echo   - webdriver-manager (for ChromeDriver)
echo.
echo ================================================
echo.

REM Upgrade pip first
echo Upgrading pip...
python -m pip install --upgrade pip
echo.

REM Install required packages
echo Installing packages...
python -m pip install pandas openpyxl selenium requests webdriver-manager

echo.
echo ================================================
echo   Installation Complete!
echo ================================================
echo.
echo You can now run: run.bat
echo.
pause

