@echo off
REM 設定代碼頁為 UTF-8
chcp 65001 >nul
echo === ExcelCode Pro - 安裝環境 ===

where python >nul 2>nul
if errorlevel 1 (
    echo [錯誤] 找不到 Python，請先安裝 Python 3.x！
    start https://www.python.org/downloads
    pause
    exit /b 1
)

echo 檢查 pip...
python -m ensurepip
python -m pip install --upgrade pip

echo 安裝必要套件中...
python -m pip install -r requirements.txt

echo [完成] 所有依賴安裝完成！
pause
