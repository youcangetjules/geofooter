@echo off
REM Uninstall GURI GUI from Windows Startup

echo Removing GURI GUI from Windows Startup...

set STARTUP_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
set SHORTCUT_NAME=GURI GUI.lnk

if exist "%STARTUP_DIR%\%SHORTCUT_NAME%" (
    del "%STARTUP_DIR%\%SHORTCUT_NAME%"
    echo Successfully removed GURI GUI from Windows Startup.
) else (
    echo GURI GUI was not found in Windows Startup.
)

pause

