@echo off
setlocal
cd /d "%~dp0"
title ExcelHandle Local Service

echo Starting ExcelHandle...
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0start.ps1"
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
    echo.
    echo ExcelHandle failed to start. See the error above or start-error.log.
    pause
)

exit /b %EXIT_CODE%
