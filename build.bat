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

REM 確認並複製 templates 資料夾到 dist 目錄
if exist templates (
    echo Copying templates folder to dist directory...
    if not exist dist\templates mkdir dist\templates
    xcopy /E /I /Y templates dist\templates > nul
) else (
    echo Warning: templates folder not found
    mkdir templates
    echo Templates folder created. Please add template files.
)

REM 確認並複製 configs 資料夾到 dist 目錄
if exist configs (
    echo Copying configs folder to dist directory...
    if not exist dist\configs mkdir dist\configs
    xcopy /E /I /Y configs dist\configs > nul
) else (
    echo Warning: configs folder not found
    mkdir configs
    echo Configs folder created. Please add config files.
)

py -3.13 -m pip freeze > requirements.txt

echo Build process completed!