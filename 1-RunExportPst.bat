@echo off
cls
echo Running PST Export via Exchange Management Shell...

REM Get the path of the batch file
set "ScriptDir=%~dp0"

REM Run the PowerShell script located in the same directory
powershell -NoProfile -ExecutionPolicy Bypass -File "%ScriptDir%Export_Pst.ps1"

echo.
echo Script finished.

REM Check for error and handle accordingly
if errorlevel 1 (
    echo Script encountered an error.
    pause
) else (
    timeout /t 10
)