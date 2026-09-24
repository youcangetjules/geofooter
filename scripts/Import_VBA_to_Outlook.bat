@echo off
REM ============================================================================
REM  Import_VBA_to_Outlook.bat
REM  Sync AES VBA modules from VBA\ into the running Outlook project.
REM  Always uses 32-bit Windows PowerShell 5.1 when available so COM can
REM  attach to 32-bit Outlook (bare "powershell" may be PowerShell 7).
REM ============================================================================
setlocal
title Import AES VBA into Outlook

for %%I in ("%~dp0..") do set "ROOT=%%~fI"

set "PS32=%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
set "PS64=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%PS32%" (
    set "PS=%PS32%"
) else if exist "%PS64%" (
    set "PS=%PS64%"
) else (
    set "PS=powershell"
)

echo.
echo Usage tips:
echo   1st time:  this bat with arg  -EnableAccessVBOM
echo              then fully quit Outlook and reopen, then run again
echo   Normal:    no args  (re-imports all MSCAN modules)
echo   Optional:  -SyncThisOutlookSession  (also overwrites built-in ThisOutlookSession)
echo.
echo Using: %PS%
echo.

"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Import_VBA_to_Outlook.ps1" -Root "%ROOT%" %*

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
