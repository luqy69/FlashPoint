@echo off
setlocal enabledelayedexpansion
REM ================================================================
REM  FlashPoint AI - PowerPoint Generator
REM  Debug Runner with Advanced Error Logging
REM ================================================================

echo.
echo ============================================
echo   FlashPoint AI - Presentation Generator
echo   Debug Mode with Error Logging
echo ============================================
echo.

REM Create logs directory
if not exist "logs" mkdir logs

REM Generate timestamped log filename
for /f "tokens=2 delims==" %%I in ('wmic os get localdatetime /format:list 2^>nul') do set datetime=%%I
set LOGFILE=logs\run_%datetime:~0,8%_%datetime:~8,6%.log

echo [*] Log file: %LOGFILE%
echo.

REM Log system info
echo ================================================================ > "%LOGFILE%"
echo FlashPoint AI - Debug Log >> "%LOGFILE%"
echo Started: %date% %time% >> "%LOGFILE%"
echo ================================================================ >> "%LOGFILE%"
echo. >> "%LOGFILE%"

echo --- SYSTEM INFO --- >> "%LOGFILE%"
echo OS: %OS% >> "%LOGFILE%"
echo ComputerName: %COMPUTERNAME% >> "%LOGFILE%"
echo UserName: %USERNAME% >> "%LOGFILE%"
echo. >> "%LOGFILE%"

REM Check Python installation
echo --- PYTHON CHECK --- >> "%LOGFILE%"
python --version >> "%LOGFILE%" 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found in PATH! >> "%LOGFILE%"
    echo [!] Python is not installed or not in PATH.
    echo     Please install Python from https://python.org
    pause
    exit /b 1
)
echo [OK] Python found >> "%LOGFILE%"
echo. >> "%LOGFILE%"

REM Check pip
pip --version >> "%LOGFILE%" 2>&1
echo. >> "%LOGFILE%"

REM Check critical dependencies
echo --- DEPENDENCY CHECK --- >> "%LOGFILE%"
set MISSING=0

for %%M in (selenium pptx colorama bs4 pyperclip) do (
    python -c "import %%M" 2>nul
    if errorlevel 1 (
        echo [MISSING] %%M >> "%LOGFILE%"
        set MISSING=1
    ) else (
        echo [OK] %%M >> "%LOGFILE%"
    )
)

if !MISSING!==1 (
    echo [!] Missing dependencies detected. Installing...
    echo [*] Installing missing dependencies... >> "%LOGFILE%"
    pip install -r requirements.txt >> "%LOGFILE%" 2>&1
    echo [+] Installation complete!
    echo. >> "%LOGFILE%"
)

REM Check Chrome
echo --- CHROME CHECK --- >> "%LOGFILE%"
where chrome >nul 2>&1
if errorlevel 1 (
    if exist "C:\Program Files\Google\Chrome\Application\chrome.exe" (
        echo [OK] Chrome found at default location >> "%LOGFILE%"
    ) else if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" (
        echo [OK] Chrome found at x86 location >> "%LOGFILE%"
    ) else (
        echo [WARNING] Chrome not found in standard locations >> "%LOGFILE%"
        echo [!] Warning: Chrome may not be installed
    )
) else (
    echo [OK] Chrome in PATH >> "%LOGFILE%"
)

REM Get Chrome version
for /f "tokens=*" %%V in ('reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version 2^>nul ^| findstr version') do (
    echo Chrome Version: %%V >> "%LOGFILE%"
)
echo. >> "%LOGFILE%"

REM Check config files
echo --- CONFIG CHECK --- >> "%LOGFILE%"
for %%F in (slide_formating_prompt prompt_lvl_1 prompt_lvl_2 prompt_lvl_3) do (
    if exist "%%F" (
        echo [OK] %%F found >> "%LOGFILE%"
    ) else if exist "config\%%F" (
        echo [OK] config\%%F found >> "%LOGFILE%"
    ) else (
        echo [MISSING] %%F >> "%LOGFILE%"
    )
)
echo. >> "%LOGFILE%"

echo --- APPLICATION START --- >> "%LOGFILE%"
echo [*] Starting at %time% >> "%LOGFILE%"
echo. >> "%LOGFILE%"

echo [*] Starting presentation generator...
echo     Console output is also being saved to %LOGFILE%
echo.

REM Run the generator with output tee'd to log file
REM Using a temporary file to capture both stdout and stderr
python ppt_generator.py 2>&1 | findstr /r "." & python ppt_generator.py >> "%LOGFILE%" 2>&1

REM Capture exit code
set EXITCODE=%ERRORLEVEL%

echo. >> "%LOGFILE%"
echo --- APPLICATION EXIT --- >> "%LOGFILE%"
echo Exit Code: %EXITCODE% >> "%LOGFILE%"
echo Ended: %date% %time% >> "%LOGFILE%"

if %EXITCODE% NEQ 0 (
    echo.
    echo ============================================
    echo [!] Program exited with error code: %EXITCODE%
    echo     Check log file for details: %LOGFILE%
    echo ============================================
    echo.
    echo [ERROR] Exit code %EXITCODE% >> "%LOGFILE%"
) else (
    echo.
    echo [+] Program completed successfully.
    echo [OK] Clean exit >> "%LOGFILE%"
)

echo.
echo Log saved to: %LOGFILE%
echo.
pause
