@echo off
setlocal enabledelayedexpansion

:: Set script and log paths
set "ScriptDir=D:\Export_PST"
set "LogDir=C:\PST_Export_Logs"

:: Create log directory if not exists
if not exist "%LogDir%" (
    mkdir "%LogDir%"
)

:: Generate timestamped log file
for /f "tokens=1-4 delims=/ " %%a in ('date /t') do (
    set "YYYY=%%d"
    set "MM=%%b"
    set "DD=%%c"
)
for /f "tokens=1-2 delims=: " %%a in ("%time%") do (
    set "HH=%%a"
    set "Min=%%b"
)
set "LogFile=%LogDir%\0-PST_Export_%YYYY%-%MM%-%DD%_%HH%-%Min%.log"

echo ===== Starting PST Export Process on %COMPUTERNAME% ===== >> "%LogFile%"
echo Log File: %LogFile%
echo.

:: Send start notification
echo ===== Sending Start Notification ===== >> "%LogFile%"
powershell -ExecutionPolicy Bypass -File "%ScriptDir%\Send-Notification.ps1" -Subject "PST Export Started" -Body "The PST export process has started on %COMPUTERNAME%." >> "%LogFile%" 2>&1

echo ===== Starting PST Export Process ===== >> "%LogFile%"
call "%ScriptDir%\1-RunExportPst.bat" >> "%LogFile%" 2>&1

echo ===== Waiting 180 seconds before Archiving... ===== >> "%LogFile%"
timeout /t 180 >> "%LogFile%" 2>&1

echo ===== Starting Archiving PST Process ===== >> "%LogFile%"
call "%ScriptDir%\2-Archeive_Data.bat" >> "%LogFile%" 2>&1

echo ===== Waiting 10 seconds before Moving Archived Data... ===== >> "%LogFile%"
timeout /t 10 >> "%LogFile%" 2>&1

echo ===== Starting Archiving PST Move to Network Drive ===== >> "%LogFile%"
call "%ScriptDir%\3-Move_Pst.bat" >> "%LogFile%" 2>&1

echo ===== Waiting 10 seconds before Clearing Export Request... ===== >> "%LogFile%"
timeout /t 10 >> "%LogFile%" 2>&1

echo ===== Removing Completed Export Requests ===== >> "%LogFile%"
call "%ScriptDir%\4-Clear_MailBox_Request.bat" >> "%LogFile%" 2>&1

echo ===== Sending Completion Notification ===== >> "%LogFile%"
powershell -ExecutionPolicy Bypass -File "%ScriptDir%\Send-Notification.ps1" -Subject "PST Export Completed" -Body "The PST export and archiving process completed successfully on %COMPUTERNAME%." >> "%LogFile%" 2>&1

echo ===== All Tasks Complete ===== >> "%LogFile%"
echo.
pause
