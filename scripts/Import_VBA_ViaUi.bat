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
  echo.
  echo If Norton / AV blocked keystroke automation:
  echo   1. Norton ^> Settings ^> Antivirus ^> Scans and Risks ^> Exclusions / Low Risks
  echo      Add folder: %ROOT%\scripts
  echo      Add folder: %ROOT%\VBA
  echo   2. Or use the Norton-safe guided importer ^(clipboard only, you click Import^):
  echo      %~dp0Import_VBA_Guided.bat
  echo.
  pause
  exit /b %ERR%
)
echo Import_VBA_ViaUi finished.
pause
endlocal
