@echo off

where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [提示] 找不到 Python，正在執行 install.bat 安裝環境...
    call install.bat

    rem 安裝完後再檢查一次
    where python >nul 2>nul
    if %errorlevel% neq 0 (
        echo [錯誤] 安裝後仍找不到 Python，請手動檢查安裝。
        pause
        exit /b 1
    )
)

python main.py

python -m pip freeze > requirements.txt