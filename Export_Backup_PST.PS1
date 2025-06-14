@echo off
echo ===== Starting PST Export Process =====
call "D:\Export_PST\1-RunExportPst.bat"

echo ===== Waiting 30 seconds before Archiving... =====
timeout /t 30
echo.
echo ===== Starting Archiving PST Process =====
call "D:\Export_PST\2-Archeive_Data.bat"

echo.
echo ===== Waiting 30 seconds before Moving Archeived Data... =====
timeout /t 30

echo.
echo ===== Starting Archiving PST Move to Network Drive =====
call "D:\Export_PST\3-Move_Pst.bat"

echo.
echo ===== Waiting 30 seconds before Clearing Export Request... =====
timeout /t 30

echo.
echo ===== Removing Completed Export Requests =====
call "D:\Export_PST\4-Clear_MailBox_Request.bat"

echo.
echo ===== All Tasks Complete =====
pause
