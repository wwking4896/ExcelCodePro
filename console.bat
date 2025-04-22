@echo off
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

REM 嘗試使用 Python 3.13 執行
py -3.13 -c "import sys; print(sys.version)" 2>nul && (
    echo Using Python 3.13
    py -3.13 console.py %CMD_ARGS%
) || (
    REM 如果 Python 3.13 不可用，使用任何可用的 Python 3.x
    echo Python 3.13 not found, trying to use any available Python 3.x
    py -3 console.py %CMD_ARGS%
)

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