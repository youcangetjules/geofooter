@echo off
REM ============================================================================
REM  Import_VBA_to_Outlook.bat
REM  Sync AES VBA modules from VBA\ into the running Outlook project.
REM ============================================================================
setlocal
title Import AES VBA into Outlook

for %%I in ("%~dp0..") do set "ROOT=%%~fI"

echo.
echo Usage tips:
echo   1st time:  this bat with arg  -EnableAccessVBOM
echo              then fully quit Outlook and reopen, then run again
echo   Normal:    no args  (re-imports all MSCAN modules)
echo   Optional:  -SyncThisOutlookSession  (also overwrites built-in ThisOutlookSession)
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Import_VBA_to_Outlook.ps1" -Root "%ROOT%" %*

set ERR=%ERRORLEVEL%
if %ERR% neq 0 (
    echo.
    echo Import failed with exit code %ERR%.
    pause
    exit /b %ERR%
)

echo.
echo Done. Compile + Save in the VBA editor, then restart Outlook if needed.
pause
endlocal
exit /b 0
