@echo off
REM Build script for CTS CMR Converter
REM This script builds the executable and creates the installer

echo ========================================
echo CTS CMR Converter - Build Script
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    pause
    exit /b 1
)

echo [1/4] Installing/updating dependencies...
pip install --upgrade pyinstaller openpyxl pdfplumber packaging Pillow
if errorlevel 1 (
    echo ERROR: Failed to install dependencies
    pause
    exit /b 1
)

echo.
echo [2/4] Building executable with PyInstaller...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"

REM Build using spec file if it exists, otherwise use command line
if exist "CTS_CMR_Converter.spec" (
    pyinstaller CTS_CMR_Converter.spec
) else (
    pyinstaller --name "CTS_CMR_Converter" --onefile --windowed --add-data "pdf_to_cmr.py;." --add-data "updater.py;." pdf_to_cmr_gui.py
)

if errorlevel 1 (
    echo ERROR: PyInstaller build failed
    pause
    exit /b 1
)

echo.
echo [3/4] Checking if executable was created...
if not exist "dist\CTS_CMR_Converter.exe" (
    echo ERROR: Executable not found in dist folder
    pause
    exit /b 1
)

echo ✓ Executable created: dist\CTS_CMR_Converter.exe
echo.

echo [4/4] Creating installer with Inno Setup...
REM Check if Inno Setup is installed
set INNO_PATH="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not exist %INNO_PATH% (
    echo WARNING: Inno Setup not found at %INNO_PATH%
    echo Please install Inno Setup from: https://jrsoftware.org/isdl.php
    echo.
    echo You can still use the executable in the dist folder.
    goto :success
)

REM Build installer
%INNO_PATH% setup.iss
if errorlevel 1 (
    echo WARNING: Installer creation failed
    echo.
    echo You can still use the executable in the dist folder.
    goto :success
)

echo ✓ Installer created in installer_output folder
echo.

:success
echo ========================================
echo BUILD SUCCESSFUL!
echo ========================================
echo.
echo Outputs:
if exist "dist\CTS_CMR_Converter.exe" (
    echo   ✓ Executable: dist\CTS_CMR_Converter.exe
)
if exist "installer_output" (
    for %%f in (installer_output\*.exe) do (
        echo   ✓ Installer: %%f
    )
)
echo.
echo You can now distribute these files!
echo.
pause
