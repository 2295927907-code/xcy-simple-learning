@echo off
setlocal
cd /d "%~dp0"
set "LAUNCHER=%~dp0start_spd_calculator.ps1"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%LAUNCHER%"

if errorlevel 1 (
    echo.
    echo Program failed to start. Please check SPD_startup_log.txt in this folder.
    pause
)
