@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0MastercamScan.ps1"
if errorlevel 1 (
  echo.
  echo Scan failed. Please send a screenshot of this window.
  pause
)
endlocal
