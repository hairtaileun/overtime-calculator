@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\audit-v32.ps1" %*
set RC=%ERRORLEVEL%
echo.
echo OVERTIME_AUDIT_EXIT=%RC%
exit /b %RC%
