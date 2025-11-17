@echo off
REM CTS PDF to CMR Converter - Installation Script for Windows
REM This script installs all required dependencies

echo ================================================
echo CTS Packing List to CMR Converter - Installation
echo ================================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed!
    echo.
    echo Please install Python 3.8 or higher from:
    echo https://www.python.org/downloads/
    echo.
    echo Make sure to check "Add Python to PATH" during installation!
    pause
    exit /b 1
)

echo Python is installed:
python --version
echo.

REM Upgrade pip
echo Upgrading pip...
python -m pip install --upgrade pip --quiet
echo.

REM Install required packages
echo Installing required packages...
echo - pdfplumber (for PDF extraction)
echo - openpyxl (for Excel manipulation)
echo.

python -m pip install pdfplumber openpyxl --quiet

if errorlevel 1 (
    echo.
    echo ERROR: Failed to install required packages
    pause
    exit /b 1
)

echo.
echo ================================================
echo Installation Complete!
echo ================================================
echo.
echo You can now use the converter in three ways:
echo.
echo 1. GUI (Easiest):
echo    Double-click: pdf_to_cmr_gui.py
echo.
echo 2. Command Line:
echo    python pdf_to_cmr.py 5523
echo.
echo 3. PowerShell:
echo    .\convert_pdf_to_cmr.ps1 -Input 5523
echo.
echo 4. Batch Processing:
echo    python batch_convert.py ./packing_lists
echo.
echo See README.md for detailed instructions
echo ================================================
echo.
pause
