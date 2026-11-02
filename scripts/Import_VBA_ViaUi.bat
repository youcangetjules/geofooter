@echo off
setlocal
REM Drive Outlook VBE File > Import when Application.VBE is blocked.
set "ROOT=%~dp0.."
set "PS32=%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS32%" set "PS32=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
"%PS32%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Import_VBA_ViaUi.ps1" -Root "%ROOT%" %*
set ERR=%ERRORLEVEL%
echo.
if %ERR% neq 0 (
  echo Import_VBA_ViaUi failed with exit %ERR%.
  pause
  exit /b %ERR%
)
echo Import_VBA_ViaUi finished.
pause
endlocal
