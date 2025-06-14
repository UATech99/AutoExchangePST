param (
    [string]$Subject = "PST Export Notification",
    [string]$Body = "Task completed."
)

# Email settings
$From = "<Your-Email-From-System-Has-TO-Send-Emails>" # AutomatedPST@example.com
$To = "<To-Whom-Emails-Should-Received>" #UserName@example.com
$SmtpServer = "<SMTP-Server-FQDN>" # smtp.example.com
#$SmtpPort = 25

Send-MailMessage -From $From -To $To -Subject $Subject -Body $Body -SmtpServer $SmtpServer
