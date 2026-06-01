@echo off
title Excel to CSV Converter - Build Tool
echo ==========================================
echo   Excel to CSV Converter - Build Tool
echo   Author: Jack Wang
echo   Blog: https://wangjinming.com
echo ==========================================
echo.

:: Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found! Please install Python 3.6+
    echo Download: https://www.python.org/downloads/
    pause
    exit /b 1
)
echo [OK] Python detected

:: Install dependencies
echo [INFO] Installing dependencies...
pip install pyinstaller openpyxl xlrd -q
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install dependencies
    pause
    exit /b 1
)
echo [OK] Dependencies installed

:: Build EXE
echo [INFO] Building EXE ...
pyinstaller --onefile --noconsole --name "ExcelToCSV" excel_to_csv.py
if %errorlevel% neq 0 (
    echo [ERROR] Build failed
    pause
    exit /b 1
)
echo [OK] Build completed!

:: Clean up temp files
echo [INFO] Cleaning up...
rmdir /s /q build >nul 2>&1
del *.spec >nul 2>&1
rmdir /s /q __pycache__ >nul 2>&1

move dist\ExcelToCSV.exe . >nul 2>&1
rmdir dist >nul 2>&1

echo.
echo ==========================================
echo   Output: ExcelToCSV.exe
echo.
echo   Copy ExcelToCSV.exe to any Windows PC
echo   and double-click to run.
echo   No Python installation required!
echo ==========================================
echo.
pause
