@echo off
echo ======================================
echo Clearing Windows Icon Cache
echo ======================================
echo.

echo [1] Killing Explorer...
taskkill /f /im explorer.exe

echo.
echo [2] Deleting icon cache files...
cd /d %userprofile%\AppData\Local
if exist IconCache.db del /f /q IconCache.db
if exist Microsoft\Windows\Explorer\iconcache* del /f /q Microsoft\Windows\Explorer\iconcache*

echo.
echo [3] Restarting Explorer...
start explorer.exe

echo.
echo ======================================
echo Icon cache cleared successfully!
echo ======================================
echo.
echo The new icon should now appear.
echo If not, try restarting your computer.
pause
