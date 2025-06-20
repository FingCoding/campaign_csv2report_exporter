@echo off
REM fam8 campaign report system dependencies installation batch
REM Location: D:\rep05\Pythonエンベッダブル検証環境\campaign_csv2report_exporter\install_dependencies.bat
REM Purpose: Install required libraries to Python embeddable environment

echo ============================================================
echo fam8 campaign report dependencies installation started
echo Installation time: %date% %time%
echo ============================================================

REM Change working directory to current batch file location
cd /d "%~dp0"

REM Python embeddable environment path setting
set PYTHON_HOME=D:\rep05\Pythonエンベッダブル検証環境\Python実行環境
set PYTHON_EXE=%PYTHON_HOME%\python.exe
set PIP_EXE=%PYTHON_HOME%\Scripts\pip.exe

REM Python executable check
if not exist "%PYTHON_EXE%" (
    echo [ERROR] Python executable not found: %PYTHON_EXE%
    echo Process aborted
    pause
    exit /b 1
)

echo [INFO] Python environment: %PYTHON_EXE%
echo [INFO] Working directory: %cd%

REM Check if pip is installed
if not exist "%PIP_EXE%" (
    echo [INFO] pip not found, installing pip...
    "%PYTHON_EXE%" "%PYTHON_HOME%\get-pip.py"
    if %ERRORLEVEL% NEQ 0 (
        echo [ERROR] pip installation failed
        pause
        exit /b 1
    )
    echo [SUCCESS] pip installation completed
)

REM Check requirements.txt existence
if not exist "requirements.txt" (
    echo [ERROR] requirements.txt not found: %cd%\requirements.txt
    echo Process aborted
    pause
    exit /b 1
)

echo [INFO] Installing required libraries from requirements.txt...

REM Install required libraries
"%PIP_EXE%" install -r requirements.txt --no-warn-script-location

REM Installation result check
if %ERRORLEVEL% EQU 0 (
    echo ============================================================
    echo [SUCCESS] All dependencies installed successfully
    echo Completion time: %date% %time%
    echo ============================================================
) else (
    echo ============================================================
    echo [ERROR] Error occurred during dependencies installation
    echo Error code: %ERRORLEVEL%
    echo Completion time: %date% %time%
    echo ============================================================
    pause
    exit /b %ERRORLEVEL%
)

echo [INFO] Verifying installed packages...
"%PIP_EXE%" list

echo ============================================================
echo Installation verification completed
echo You can now run: run_fam8_report.bat
echo ============================================================

pause
exit /b 0