@echo off
REM GURI Database Viewer Launcher
REM Launches the GURI GUI application

echo ========================================
echo   GURI Database Viewer ^& Manager
echo ========================================
echo.

REM Prefer the venv created by Install_GeoFooter.bat
set "PY=%~dp0..\.venv\Scripts\python.exe"
if not exist "%PY%" set "PY=python"

"%PY%" --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed. Run scripts\Install_GeoFooter.bat first.
    pause
    exit /b 1
)

echo Starting GURI GUI...
echo.

"%PY%" "%~dp0..\guri\gui.py"

if errorlevel 1 (
    echo.
    echo ERROR: Failed to launch GURI GUI
    echo Check the error messages above
    pause
)

