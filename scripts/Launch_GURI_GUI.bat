@echo off
REM GURI Database Viewer Launcher
REM Launches the GURI GUI application

echo ========================================
echo   GURI Database Viewer ^& Manager
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.7+ from https://python.org
    pause
    exit /b 1
)

echo Starting GURI GUI...
echo.

REM Run the GUI
python "%~dp0..\guri\gui.py"

if errorlevel 1 (
    echo.
    echo ERROR: Failed to launch GURI GUI
    echo Check the error messages above
    pause
)

