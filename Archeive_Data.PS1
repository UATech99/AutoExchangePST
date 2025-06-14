
# Define source and destination paths
$sourcePath = "C:\pst"
$currentDate = (Get-Date).ToString("dd.MM.yyyy")
$currentTime = (Get-Date).ToString("HHmmss")

# Ensure the destination folder exists
$sourceFolder = "C:\pst"
$rarExe = "C:\Program Files\WinRAR\rar.exe"  # Adjust if WinRAR is in a different location

# Check if WinRAR exists
if (-not (Test-Path $rarExe)) {
    Write-Host "❌ WinRAR not found at $rarExe. Please check the path." -ForegroundColor Red
    exit
}

# Get all .pst files in the source folder
$pstFiles = Get-ChildItem -Path $sourceFolder -Filter *.pst

foreach ($file in $pstFiles) {
    $fileName = $file.BaseName
    $sourceFile = $file.FullName
    $rarFile = Join-Path $sourceFolder "$fileName.rar"

    # Create RAR archive
    & "$rarExe" a -ep1 "$rarFile" "$sourceFile" | Out-Null

    # Check if RAR was successfully created, then delete original PST
    if (Test-Path $rarFile) {
        Remove-Item -Path $sourceFile -Force
        Write-Host "✅ Compressed and deleted: $($file.Name)" -ForegroundColor Green
    } else {
        Write-Host "❌ Failed to compress: $($file.Name)" -ForegroundColor Red
    }
}

Write-Host "`n🎉 All .pst files have been processed." -ForegroundColor Cyan
