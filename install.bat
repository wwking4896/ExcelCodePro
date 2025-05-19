@echo off
REM 設定代碼頁為 UTF-8
chcp 65001 >nul

echo ========================================
echo ExcelCode Pro Dependency Installation
echo ========================================

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: Python not found
    echo Please install Python 3.x first
    echo Download: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo Detected Python version:
python --version

echo.
echo Installing dependencies...
pip install -r requirements.txt

if %errorlevel% equ 0 (
    echo.
    echo ========================================
    echo Installation completed successfully!
    echo You can now run: python main.py
    echo ========================================
) else (
    echo.
    echo ========================================
    echo Error occurred during installation
    echo Please check your network connection
    echo ========================================
)

pause