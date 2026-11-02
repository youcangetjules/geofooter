@echo off
setlocal
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Backup_VbaProject_OTM.ps1" %*
if errorlevel 1 pause
endlocal
