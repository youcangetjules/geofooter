@echo off
setlocal
set "ROOT=%~dp0.."
set "PS32=%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS32%" set "PS32=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
"%PS32%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Import_VBA_Guided.ps1" -Root "%ROOT%" %*
set ERR=%ERRORLEVEL%
if %ERR% neq 0 pause
endlocal
exit /b %ERR%
