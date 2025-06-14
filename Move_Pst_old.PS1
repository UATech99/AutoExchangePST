$sourceFolder = "C:\pst"
$destinationRoot = "\\192.168.4.54\fileserver\backup\Daily Backup Copy\PST-Backup\June 2025"  # <-- Replace with your actual network share
$dateFolder = Get-Date -Format 'dd-MM-yyyy'
$destinationPath = Join-Path $destinationRoot $dateFolder
$logfolder =  "c:\PST_Export_Logs"
$logFile = Join-Path $logfolder "move_log.txt"      # Move_Pst
# Start log
Add-Content -Path $logFile -Value "`n=== RAR Move Log: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="

# Check if destination exists, if not create
if (-not (Test-Path $destinationPath)) {
    New-Item -ItemType Directory -Path $destinationPath | Out-Null
    Add-Content -Path $logFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  Created folder: $destinationPath"
}

# Get all .rar files in source
$rarFiles = Get-ChildItem -Path $sourceFolder -Filter *.rar

foreach ($file in $rarFiles) {
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $destinationFile = Join-Path $destinationPath $file.Name

    try {
        Move-Item -Path $file.FullName -Destination $destinationFile -Force
        $msg = "$timestamp Moved: $($file.Name) → $destinationPath"
        Write-Host $msg -ForegroundColor Green
    } catch {
        $msg = "$timestamp  Failed to move: $($file.Name) — $($_.Exception.Message)"
        Write-Host $msg -ForegroundColor Red
    }

    Add-Content -Path $logFile -Value $msg
}

# Final log entry
$finalMsg = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') 🎉 All .rar files processed."
Write-Host "`n$finalMsg" -ForegroundColor Cyan
Add-Content -Path $logFile -Value $finalMsg
