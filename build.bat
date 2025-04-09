@echo off
REM 嘗試使用 Python 3.13 執行
py -3.13 -c "import sys; print(sys.version)" 2>nul && (
    echo Using Python 3.13
    py -3.13 build.py
) || (
    REM 如果 Python 3.13 不可用，使用任何可用的 Python 3.x
    echo Python 3.13 not found, trying to use any available Python 3.x
    py -3 build.py
)

REM 複製 README.md 到 dist 目錄
copy README.md dist\ > nul

py -3.13 -m pip freeze > requirements.txt