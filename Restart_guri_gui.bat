@echo off
REM ============================================================================
REM  Restart_guri_gui.bat
REM  Kill stuck GeoFooter / AES / GURI Python processes that can prevent
REM  guri_gui.py from starting, then relaunch GURI GUI.
REM ============================================================================
setlocal
title Restart GURI GUI

REM %~dp0 always ends with \, which escapes the closing " when passed as
REM -Root "C:\GeoFooter\" — strip it before quoting.
set "ROOT=%~dp0"
if "%ROOT:~-1%"=="\" set "ROOT=%ROOT:~0,-1%"

powershell -NoProfile -ExecutionPolicy Bypass -File "%ROOT%\Restart_guri_gui.ps1" -Root "%ROOT%"

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
