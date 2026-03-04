@echo off
setlocal enabledelayedexpansion
REM ================================================================
REM  FlashPoint AI - Debug Runner
REM  Runs the executable with error logging
REM ================================================================

echo.
echo ============================================
echo   FlashPoint AI - Debug Mode
echo ============================================
echo.

REM Create logs directory
if not exist "logs" mkdir logs

REM Generate timestamped log filename
for /f "tokens=2 delims==" %%I in ('wmic os get localdatetime /format:list 2^>nul') do set datetime=%%I
set LOGFILE=logs\debug_%datetime:~0,8%_%datetime:~8,6%.log

echo [*] Debug log: %LOGFILE%
echo.

REM Log header
echo ================================================================ > "%LOGFILE%"
echo FlashPoint AI - Debug Log >> "%LOGFILE%"
echo Started: %date% %time% >> "%LOGFILE%"
echo ================================================================ >> "%LOGFILE%"
echo. >> "%LOGFILE%"

REM System info
echo OS: %OS% >> "%LOGFILE%"
echo Computer: %COMPUTERNAME% >> "%LOGFILE%"
echo. >> "%LOGFILE%"

REM Chrome check
echo --- CHROME --- >> "%LOGFILE%"
for /f "tokens=*" %%V in ('reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version 2^>nul ^| findstr version') do (
    echo Chrome: %%V >> "%LOGFILE%"
    echo [+] Chrome: %%V
)
echo. >> "%LOGFILE%"

REM Check exe exists
if not exist "FlashPoint_AI.exe" (
    echo [!] FlashPoint_AI.exe not found!
    echo [ERROR] FlashPoint_AI.exe not found >> "%LOGFILE%"
    pause
    exit /b 1
)

REM Check config
echo --- CONFIG --- >> "%LOGFILE%"
if exist "config" (
    dir /b config >> "%LOGFILE%" 2>&1
    echo [OK] Config folder found >> "%LOGFILE%"
) else (
    echo [WARNING] No config folder >> "%LOGFILE%"
)
echo. >> "%LOGFILE%"

echo --- RUN START --- >> "%LOGFILE%"
echo [*] Running FlashPoint_AI.exe ...
echo.

REM Run exe and capture output to log
FlashPoint_AI.exe >> "%LOGFILE%" 2>&1
set EXITCODE=%ERRORLEVEL%

echo. >> "%LOGFILE%"
echo --- RUN END --- >> "%LOGFILE%"
echo Exit Code: %EXITCODE% >> "%LOGFILE%"
echo Ended: %date% %time% >> "%LOGFILE%"

if %EXITCODE% NEQ 0 (
    echo.
    echo ==========================================
    echo [!] Error detected (code: %EXITCODE%)
    echo     Full log: %LOGFILE%
    echo ==========================================
) else (
    echo.
    echo [+] Completed successfully.
)

echo.
echo Debug log saved: %LOGFILE%
echo.
pause
