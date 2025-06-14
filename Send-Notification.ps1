param (
    [string]$Subject = "PST Export Notification",
    [string]$Body = "Task completed."
)

# Email settings
$From = "EDLAutomation@edl2.com"
$To = "umar@edl2.com"
$SmtpServer = "Exchangeedl2.edl2.com"
#$SmtpPort = 25

Send-MailMessage -From $From -To $To -Subject $Subject -Body $Body -SmtpServer $SmtpServer
