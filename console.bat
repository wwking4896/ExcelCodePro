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

REM ==========================================================
REM ExcelCode Pro - Console Mode Batch File
REM ==========================================================

REM 設定默認參數
set CONFIG_FILE=config.json
set OUTPUT_FILE=output.h
set VERBOSE=0

REM 解析命令行參數
:parse_args
if "%~1"=="" goto :continue
if /i "%~1"=="--config" (
    set CONFIG_FILE=%~2
    shift
    shift
    goto :parse_args
)
if /i "%~1"=="-c" (
    set CONFIG_FILE=%~2
    shift
    shift
    goto :parse_args
)
if /i "%~1"=="--output" (
    set OUTPUT_FILE=%~2
    shift
    shift
    goto :parse_args
)
if /i "%~1"=="-o" (
    set OUTPUT_FILE=%~2
    shift
    shift
    goto :parse_args
)
if /i "%~1"=="--verbose" (
    set VERBOSE=1
    shift
    goto :parse_args
)
if /i "%~1"=="-v" (
    set VERBOSE=1
    shift
    goto :parse_args
)
echo Unknown parameter: %~1
shift
goto :parse_args

:continue
REM 顯示執行資訊
echo ==========================================================
echo ExcelCode Pro - Console Mode
echo ==========================================================
echo Config file: %CONFIG_FILE%
echo Output file: %OUTPUT_FILE%
if %VERBOSE%==1 (
    echo Verbose mode: Enabled
) else (
    echo Verbose mode: Disabled
)
echo ==========================================================

REM 檢查配置文件是否存在
if not exist "%CONFIG_FILE%" (
    echo Error: Config file does not exist: %CONFIG_FILE%
    goto :end
)

REM 建立命令參數
set CMD_ARGS=--config "%CONFIG_FILE%" --output "%OUTPUT_FILE%"
if %VERBOSE%==1 (
    set CMD_ARGS=%CMD_ARGS% --verbose
)

python console.py %CMD_ARGS%

:end
REM 檢查生成的輸出文件是否存在
if exist "%OUTPUT_FILE%" (
    echo Code generation successful! Output saved to: %OUTPUT_FILE%
) else (
    echo WARNING: Output file not found. Check for errors above.
)
echo ==========================================================
echo Operation completed
echo ==========================================================

REM 等待按鍵再關閉視窗（方便查看結果）
pause