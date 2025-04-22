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

REM 生成命令行版本執行檔案
echo Creating console version executable...
py -3.13 -m PyInstaller --clean --onefile --name="ExcelCode-Console" console.py

REM 將命令行版本的執行檔案複製到 dist 目錄
echo Copying console version to the distribution folder...
copy dist\ExcelCode-Console.exe dist\ExcelCode\ > nul

REM 複製 console.bat 到 dist 目錄作為使用範例
if exist console.bat (
    echo Copying console.bat example to the distribution folder...
    copy console.bat dist\ExcelCode\ > nul
)

REM 建立命令行範例配置檔案
echo Creating example config.json file...
if not exist dist\ExcelCode\examples mkdir dist\ExcelCode\examples
echo {> dist\ExcelCode\examples\config.json
echo   "excel_files": ["path/to/your/excel/file.xlsx"],>> dist\ExcelCode\examples\config.json
echo   "selected_sheet": "Sheet1",>> dist\ExcelCode\examples\config.json
echo   "selected_ranges": [>> dist\ExcelCode\examples\config.json
echo     {>> dist\ExcelCode\examples\config.json
echo       "start_row": 0,>> dist\ExcelCode\examples\config.json
echo       "start_col": 0,>> dist\ExcelCode\examples\config.json
echo       "end_row": 10,>> dist\ExcelCode\examples\config.json
echo       "end_col": 3,>> dist\ExcelCode\examples\config.json
echo       "range_str": "A1:D11">> dist\ExcelCode\examples\config.json
echo     }>> dist\ExcelCode\examples\config.json
echo   ],>> dist\ExcelCode\examples\config.json
echo   "template_type": "preset",>> dist\ExcelCode\examples\config.json
echo   "preset_template": "二維陣列">> dist\ExcelCode\examples\config.json
echo }>> dist\ExcelCode\examples\config.json

REM 建立命令行使用說明
echo Creating console usage instructions...
echo Usage:> dist\ExcelCode\console_usage.txt
echo.>> dist\ExcelCode\console_usage.txt
echo 1. Command line usage:>> dist\ExcelCode\console_usage.txt
echo    ExcelCode-Console.exe --config path/to/config.json --output path/to/output.c>> dist\ExcelCode\console_usage.txt
echo.>> dist\ExcelCode\console_usage.txt
echo 2. Batch file usage:>> dist\ExcelCode\console_usage.txt
echo    Edit console.bat and configure the parameters as needed>> dist\ExcelCode\console_usage.txt
echo    Then run console.bat>> dist\ExcelCode\console_usage.txt
echo.>> dist\ExcelCode\console_usage.txt
echo 3. Example config.json can be found in the 'examples' folder>> dist\ExcelCode\console_usage.txt

py -3.13 -m pip freeze > requirements.txt

echo Build process completed!
echo.
echo ==========================================================
echo Distribution folder contains both GUI and Console versions:
echo - GUI version: ExcelCode Pro.exe
echo - Console version: ExcelCode-Console.exe
echo ==========================================================
echo.

pause