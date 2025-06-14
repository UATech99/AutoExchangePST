try {
    # Load Exchange snap-in (for Exchange 2010/2013, if not already loaded)
    if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.* -ErrorAction SilentlyContinue)) {
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction Stop
    }
} catch {
    Write-Host "ERROR: Could not load Exchange snap-in. Are you running this on the Exchange server?" -ForegroundColor Red
    exit 1
}

Write-Host "Removing all completed mailbox export requests..." -ForegroundColor Cyan

# Remove all completed export requests without confirmation
Get-MailboxExportRequest -Status Completed | Remove-MailboxExportRequest -Confirm:$false

Write-Host "Completed export requests have been removed." -ForegroundColor Green
