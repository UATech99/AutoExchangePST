# AutoExchangePST
Automated Scripts for Exchange 2010 PST 
====================================================
Welcome to PST Export Application
About
Scripts are Developed with PowerShell and bat scripting!!
Developed by Umar Adil 
====================================================
Change following files and fields to fit it in your server.
1- Change PST Export Path in Export_Pst.PS1 in my case its c:/pst
2- Change Network Shared Folder Address according to yours in "Move_PST.PS1" Scripts.
3- To Setup Notifications Change Server, Sending/receiving Emails address in "Send_Notifications".

====================================================
How to Run this Script
====================================================
i suggest to check the process by running scripts in following orders manually.
1-RunExportPST.bat
2-Archeive_Data.bat
3-Move_PST.bat 
4-Clear_Mailbox_request.bat
====================================================
once above process is smooth you may schedual following one script which will call other scripts automatically.
0-Export_Backup_PST.bat
====================================================

Thank you !! for any query please contact umaradil99@gmail.com
