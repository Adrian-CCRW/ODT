@echo off

PowerShell.exe -ExecutionPolicy Bypass -File "Uninstall Office.ps1"

if %ERRORLEVEL% NEQ 0 (
    echo "Uninstall Office.ps1 failed. Aborting script."
    pause
    exit /b %ERRORLEVEL%
)

PowerShell.exe -ExecutionPolicy Bypass -File "script.ps1"
pause
