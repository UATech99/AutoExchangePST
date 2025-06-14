# Try to load Exchange snap-in
try {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction Stop
} catch {
    Write-Host "Failed to load Exchange snap-in. Are you running this on the Exchange server?"
    exit 1
}

# Setup logging
$LogPath = "C:\PST_Export_Logs"
$LogFile = Join-Path $LogPath "ExportLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
if (!(Test-Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force
}
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$timestamp - $Message"
    # $line | Tee-Object -FilePath $LogFile -Append ## Enable when Latest Powershell is installed
    Write-Host $line
    Add-Content -Path $LogFile -Value $line
}

Write-Log "Starting PST export process..."

# Define export path

$ServerName = $env:COMPUTERNAME
$PstShareName = "pst" 
$ExportPath = "\\$ServerName\$PstShareName"
$Mailboxes = Get-Mailbox

foreach ($mbx in $Mailboxes) {
    $alias = $mbx.Alias
    $filePath = Join-Path $ExportPath "$alias.pst"
    try {
        New-MailboxExportRequest -Mailbox $mbx.Identity -FilePath $filePath
        Write-Log "Started export for $alias to $filePath"
    } catch {
        Write-Log "ERROR: Failed to start export for $alias - $_"
    }
}

Write-Log "Export requests submitted. Now monitoring status..."

# Monitor progress
do {
    $inProgress = Get-MailboxExportRequest | Where-Object {$_.Status -eq "InProgress"}
    if ($inProgress) {
        Write-Host "`n--- Exports In Progress ---" -ForegroundColor Yellow
        $inProgress | ForEach-Object {
            Write-Host ("{0} - {1}" -f $_.Mailbox, $_.StatusDetail)
        }
    } else {
        Write-Host "`nAll exports complete or queued."
    }
    Start-Sleep -Seconds 30
} while ($inProgress)

Write-Log "Export process complete."
