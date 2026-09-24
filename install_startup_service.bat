@echo off
REM Install GURI GUI as Windows Startup Service
REM This script adds GURI GUI to Windows startup

echo Installing GURI GUI to Windows Startup...

REM Get the current script directory
set SCRIPT_DIR=%~dp0
set STARTUP_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup

REM Create shortcut in Startup folder
set SHORTCUT_NAME=GURI GUI.lnk
set TARGET_PATH=%SCRIPT_DIR%start_guri_gui.bat
set WORKING_DIR=%SCRIPT_DIR%

REM Use PowerShell to create shortcut
powershell -Command "$WshShell = New-Object -ComObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%STARTUP_DIR%\%SHORTCUT_NAME%'); $Shortcut.TargetPath = '%TARGET_PATH%'; $Shortcut.WorkingDirectory = '%WORKING_DIR%'; $Shortcut.Description = 'GURI Database Viewer and Manager'; $Shortcut.Save()"

if exist "%STARTUP_DIR%\%SHORTCUT_NAME%" (
    echo Successfully installed GURI GUI to Windows Startup!
    echo The application will start automatically when you log in.
) else (
    echo Failed to create startup shortcut.
    echo Please run this script as Administrator.
)

pause

