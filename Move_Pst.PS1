# Set variables
$sourceFolder = "C:\pst"

# Get current month and year (e.g., "June 2025")
$monthFolder = Get-Date -Format "MMMM yyyy"

# Root network backup path
$destinationRoot = "\\192.168.4.54\fileserver\backup\Daily Backup Copy\PST-Backup\$monthFolder"

# Create date-based subfolder (e.g., 29-05-2025)
$dateFolder = Get-Date -Format 'dd-MM-yyyy'
$destinationPath = Join-Path $destinationRoot $dateFolder

# Ensure log folder exists
$logfolder = "C:\PST_Export_Logs"
if (-not (Test-Path $logfolder)) {
    New-Item -ItemType Directory -Path $logfolder | Out-Null
}

# Create daily log file with timestamp
$logFile = Join-Path $logfolder ("move_log_" + (Get-Date -Format "yyyy-MM-dd_HH-mm-ss") + ".txt")

# Start log
Add-Content -Path $logFile -Value "`n=== RAR Move Log: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="

# Ensure month/year folder exists
if (-not (Test-Path $destinationRoot)) {
    New-Item -ItemType Directory -Path $destinationRoot | Out-Null
    Add-Content -Path $logFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Created folder: $destinationRoot"
}

# Ensure date-based folder exists
if (-not (Test-Path $destinationPath)) {
    New-Item -ItemType Directory -Path $destinationPath | Out-Null
    Add-Content -Path $logFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Created folder: $destinationPath"
}

# Get .rar files in source folder
$rarFiles = Get-ChildItem -Path $sourceFolder -Filter *.rar

foreach ($file in $rarFiles) {
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $destinationFile = Join-Path $destinationPath $file.Name

    try {
        Move-Item -Path $file.FullName -Destination $destinationFile -Force
        $msg = "$timestamp Moved: $($file.Name) → $destinationPath"
        Write-Host $msg -ForegroundColor Green
    } catch {
        $msg = "$timestamp Failed to move: $($file.Name) — $($_.Exception.Message)"
        Write-Host $msg -ForegroundColor Red
    }

    Add-Content -Path $logFile -Value $msg
}

# Final message
$finalMsg = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') 🎉 All .rar files processed."
Write-Host "`n$finalMsg" -ForegroundColor Cyan
Add-Content -Path $logFile -Value $finalMsg
