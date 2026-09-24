@echo off
REM ============================================================================
REM  Restart_guri_gui.bat
REM  Kill stuck GeoFooter / AES / GURI Python processes that can prevent
REM  guri\gui.py from starting, then relaunch GURI GUI.
REM ============================================================================
setlocal
title Restart GURI GUI

REM Suite root is the parent of scripts\. Resolve it without a trailing \,
REM which would escape the closing " in -Root "...".
for %%I in ("%~dp0..") do set "ROOT=%%~fI"

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Restart_guri_gui.ps1" -Root "%ROOT%"

if errorlevel 1 (
    echo.
    echo Restart failed.
    pause
    exit /b 1
)

echo.
echo Done. You can close this window.
timeout /t 3 /nobreak >nul
endlocal
exit /b 0
