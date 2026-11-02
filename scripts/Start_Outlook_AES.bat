@echo off
setlocal
REM Restore VbaProject.OTM from backup if missing, then start Outlook.
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Ensure_VbaProject_OTM.ps1"
if errorlevel 1 (
  echo Ensure_VbaProject_OTM failed — start Outlook anyway? Closing in 8s if you ignore.
  timeout /t 8 >nul
)
start "" "%ProgramFiles(x86)%\Microsoft Office\root\Office16\OUTLOOK.EXE"
if errorlevel 1 start "" "%ProgramFiles%\Microsoft Office\root\Office16\OUTLOOK.EXE"
endlocal
