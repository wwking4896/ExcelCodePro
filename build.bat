@echo off
REM 設定代碼頁為 UTF-8
chcp 65001 >nul
REM Enhanced build script with better error handling
setlocal enabledelayedexpansion

echo ========================================
echo   ExcelCode Pro Enhanced Build Script
echo ========================================

REM Check Python installation
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] 找不到 Python，正在執行 install.bat 安裝環境...
    call install.bat

    rem Check again after installation
    where python >nul 2>nul
    if %errorlevel% neq 0 (
        echo [ERROR] 安裝後仍找不到 Python，請手動檢查安裝。
        pause
        exit /b 1
    )
)

echo [INFO] Python 環境檢查完成

REM Execute enhanced build script
echo [INFO] 執行增強版建置腳本...
python build.py

REM Check build result
if %errorlevel% equ 0 (
    echo.
    echo ========================================
    echo          BUILD COMPLETED SUCCESSFULLY!
    echo ========================================
    echo.
    echo 建置檔案位置: dist/
    echo GUI 版本: dist/ExcelCode Pro.exe
    echo Console 版本: dist/ExcelCode-Console.exe
    echo 完整套件: dist/ExcelCode/
    echo.
) else (
    echo.
    echo ========================================
    echo             BUILD FAILED!
    echo ========================================
    echo.
    echo 請檢查上方錯誤訊息並修正問題後重新建置
    echo.
)

echo 按任意鍵繼續...
pause >nul