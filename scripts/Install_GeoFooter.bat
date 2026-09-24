@echo off
REM Master installer for GeoFooter / AES / GURI / Aura.
REM PostgreSQL is required. This launcher only starts the PowerShell installer.
setlocal
title Install GeoFooter
for %%I in ("%~dp0..") do set "ROOT=%%~fI"

echo.
echo GeoFooter needs PostgreSQL before GURI can store records.
echo The installer will check for it and tell you what to install if it is missing.
echo.

set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS%" set "PS=powershell"

"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-GeoFooter.ps1" -Root "%ROOT%" %*
set ERR=%ERRORLEVEL%
if %ERR% neq 0 (
    echo.
    echo Installer exited with code %ERR%.
    pause
    exit /b %ERR%
)
echo.
pause
exit /b 0
