@echo off
REM Install GURI GUI as Windows Service
REM Requires pywin32 to be installed: pip install pywin32

echo Installing GURI GUI as Windows Service...
echo.

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python is not installed or not in PATH
    pause
    exit /b 1
)

REM Check if pywin32 is installed
python -c "import win32serviceutil" >nul 2>&1
if errorlevel 1 (
    echo pywin32 is not installed. Installing...
    pip install pywin32
    if errorlevel 1 (
        echo Failed to install pywin32
        pause
        exit /b 1
    )
)

REM Install the service
echo Installing service...
python guri_gui_service.py install

if errorlevel 1 (
    echo Failed to install service. Please run as Administrator.
    pause
    exit /b 1
)

echo.
echo Service installed successfully!
echo.
echo To start the service, run as Administrator:
echo   python guri_gui_service.py start
echo.
echo Or use Services.msc to start it.
echo.
pause

