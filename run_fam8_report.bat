@echo off
REM fam8 campaign report auto aggregation system execution batch
REM Location: D:\rep05\PythonŠÂ‹«\campaign_csv2report_exporter\run_fam8_report.bat
REM Execution time: Daily 09:15-09:30 (executed from task scheduler)

echo ============================================================
echo fam8 campaign CSV processing started
echo Execution time: %date% %time%
echo ============================================================

REM Change working directory to current batch file location
cd /d "%~dp0"

REM Python embeddable environment path setting (FIXED PATH)
set PYTHON_HOME=D:\rep05\PythonŠÂ‹«\PythonŠÂ‹«
set PYTHON_EXE=%PYTHON_HOME%\python.exe

REM Alternative Python paths (fallback)
if not exist "%PYTHON_EXE%" (
    set PYTHON_HOME=D:\rep05\PythonŠÂ‹«
    set PYTHON_EXE=%PYTHON_HOME%\python.exe
)

if not exist "%PYTHON_EXE%" (
    set PYTHON_HOME=D:\rep05\PythonŠÂ‹«\campaign_csv2report_exporter
    set PYTHON_EXE=%PYTHON_HOME%\python.exe
)

REM System Python fallback
if not exist "%PYTHON_EXE%" (
    set PYTHON_EXE=python.exe
    python.exe --version >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] Python executable not found in any location
        echo Checked paths:
        echo - D:\rep05\PythonŠÂ‹«\PythonŠÂ‹«\python.exe
        echo - D:\rep05\PythonŠÂ‹«\python.exe
        echo - D:\rep05\PythonŠÂ‹«\campaign_csv2report_exporter\python.exe
        echo - System PATH python.exe
        echo Process aborted
        pause
        exit /b 1
    )
)

REM main.py existence check
if not exist "main.py" (
    echo [ERROR] main.py not found: %cd%\main.py
    echo Current directory contents:
    dir *.py
    echo Process aborted
    pause
    exit /b 1
)

REM config.toml existence check
if not exist "config.toml" (
    echo [ERROR] config.toml not found: %cd%\config.toml
    echo Current directory contents:
    dir *.toml
    echo Process aborted
    pause
    exit /b 1
)

REM Environment information display
echo [INFO] Python environment: %PYTHON_EXE%
echo [INFO] Working directory: %cd%
echo [INFO] Execution file: main.py

REM Python version check
echo [INFO] Python version check...
"%PYTHON_EXE%" --version
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python version check failed
    pause
    exit /b 1
)

REM Required modules check
echo [INFO] Required modules check...
"%PYTHON_EXE%" -c "import pandas, xlwings, loguru, typer, tomli; print('[SUCCESS] All required modules are available')"
if %ERRORLEVEL% NEQ 0 (
    echo [WARNING] Some required modules may be missing
    echo [INFO] Attempting to install missing modules...
    "%PYTHON_EXE%" -m pip install pandas xlwings loguru typer tomli
)

REM Execute fam8 campaign report auto aggregation (previous day)
echo [INFO] fam8 campaign report processing started...
echo [INFO] Processing target: Previous day (auto-calculated)
echo ============================================================

"%PYTHON_EXE%" main.py

REM Execution result check
set PROCESS_EXIT_CODE=%ERRORLEVEL%

if %PROCESS_EXIT_CODE% EQU 0 (
    echo ============================================================
    echo [SUCCESS] fam8 campaign CSV processing completed successfully
    echo Completion time: %date% %time%
    echo ============================================================
) else (
    echo ============================================================
    echo [ERROR] Error occurred in fam8 campaign CSV processing
    echo Error code: %PROCESS_EXIT_CODE%
    echo Completion time: %date% %time%
    echo ============================================================
    
    REM Error diagnosis
    echo [INFO] Error diagnosis:
    if %PROCESS_EXIT_CODE% EQU 1 (
        echo - General application error or invalid date format
    ) else if %PROCESS_EXIT_CODE% EQU 130 (
        echo - Process interrupted by user (Ctrl+C)
    ) else (
        echo - Unknown error code: %PROCESS_EXIT_CODE%
    )
    
    echo [INFO] Check log files in the 'log' directory for detailed error information
    if exist "log" (
        echo [INFO] Recent log files:
        dir /b /o-d log\*\*.log | more
    )
    
    REM Do not pause when executed from task scheduler
    if "%1"=="--no-pause" goto :error_end
    pause
    
    :error_end
    exit /b %PROCESS_EXIT_CODE%
)

REM Do not pause when executed from task scheduler
if "%1"=="--no-pause" goto :success_end
pause

:success_end
exit /b 0

pause