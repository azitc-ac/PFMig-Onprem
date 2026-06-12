<#
    Batch migration script for Level-1 public folders
    Migrates all immediate children of root (\) using incremental copy (delta runs)
    Uses identical naming rules as single-folder migration

    Usage:
        .\Migrate-PublicFolders-BatchLevel1.ps1
        .\Migrate-PublicFolders-BatchLevel1.ps1 -Url 'https://sv-exch02/EWS/Exchange.asmx'
        .\Migrate-PublicFolders-BatchLevel1.ps1 -WhatIf
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [string]$Url,
    [string]$AutodiscoverUrl,
    [string]$OUforMBX
)

# Source migration script
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$migrationScript = Join-Path $scriptRoot 'Migrate-PublicFolders.ps1'

if (-not (Test-Path $migrationScript)) {
    Write-Error "Migration script not found: $migrationScript"
    exit 1
}

# Function to generate target mailbox name (same logic as Migrate-PublicFolders.ps1)
function Get-TargetMailboxName {
    param([string]$PublicFolderPath)

    $last = ($PublicFolderPath.Split("\") | Select-Object -Last 1)
    if ([string]::IsNullOrWhiteSpace($last)) { $last = "ROOT" }

    $name = ("PF_{0}" -f $last) `
        -ireplace 'ä','ae' -ireplace 'ö','oe' -ireplace 'ü','ue' -ireplace 'ß','ss' `
        -ireplace ' ', '' -ireplace '\\','_'

    return $name
}

# Get all Level-1 folders (direct children of root)
Write-Host "`nFetching Level-1 public folders..." -ForegroundColor Green
$level1Folders = @(Get-PublicFolder |
    Where-Object { $_.ParentPath -eq "\" } |
    Select-Object Name, ParentPath)

if ($level1Folders.Count -eq 0) {
    Write-Host "No Level-1 folders found."
    exit 0
}

Write-Host "Found $($level1Folders.Count) Level-1 folder(s)`n" -ForegroundColor Green

# Pre-check: determine which folders are mail-enabled
Write-Host "Checking mail-enabled status..." -ForegroundColor Yellow
$folderStatus = @()
foreach ($folder in $level1Folders) {
    $sourcePath = "\$($folder.Name)"
    $isMailEnabled = $false

    try {
        $mailpf = Get-MailPublicFolder -Identity $sourcePath -ErrorAction Stop
        $isMailEnabled = $true
    } catch {
        # Not mail-enabled or error
    }

    $folderStatus += [pscustomobject]@{
        Name = $folder.Name
        Path = $sourcePath
        MailEnabled = $isMailEnabled
    }
}

Write-Host ""
Write-Host "Folder Status:" -ForegroundColor Cyan
foreach ($fs in $folderStatus) {
    $mailIcon = if ($fs.MailEnabled) { "[✓ mail-enabled]" } else { "[ ]" }
    Write-Host ("  {0,-40} {1}" -f $fs.Path, $mailIcon) -ForegroundColor $(if ($fs.MailEnabled) { 'Green' } else { 'Gray' })
}
Write-Host ""

$successCount = 0
$failCount = 0

# Batch migrate each Level-1 folder
foreach ($folder in $level1Folders) {
    $folderName = $folder.Name
    $sourcePath = "\$folderName"
    $targetMailbox = Get-TargetMailboxName -PublicFolderPath $sourcePath

    # Get mail-enabled status for this folder
    $status = $folderStatus | Where-Object { $_.Name -eq $folderName } | Select-Object -First 1
    $mailInfo = if ($status.MailEnabled) { " [MAIL-ENABLED]" } else { "" }

    Write-Host ("=" * 80) -ForegroundColor Cyan
    Write-Host ("Batch: {0} → {1}{2}" -f $sourcePath, $targetMailbox, $mailInfo) -ForegroundColor $(if ($status.MailEnabled) { 'Green' } else { 'Cyan' })
    Write-Host ("=" * 80) -ForegroundColor Cyan

    # Build parameters for child script
    $params = @{
        PublicFolderPath = $sourcePath
        TargetMailboxName = $targetMailbox
    }

    if ($Url) { $params['Url'] = $Url }
    if ($AutodiscoverUrl) { $params['AutodiscoverUrl'] = $AutodiscoverUrl }
    if ($OUforMBX) { $params['OUforMBX'] = $OUforMBX }
    if ($PSCmdlet.ShouldProcess($sourcePath)) {
        $params['Confirm'] = $false
    }

    # Call migration script
    try {
        & $migrationScript @params
        $successCount++
    } catch {
        Write-Host ("Error migrating {0}: {1}" -f $sourcePath, $_.Exception.Message) -ForegroundColor Red
        $failCount++
    }

    Write-Host ""
}

# Summary
Write-Host "`n" -ForegroundColor Green
Write-Host ("=" * 80) -ForegroundColor Green
Write-Host "Batch Migration Summary" -ForegroundColor Green
Write-Host ("=" * 80) -ForegroundColor Green
Write-Host ("  Successful   : {0}" -f $successCount) -ForegroundColor $(if ($successCount -eq $level1Folders.Count) { 'Green' } else { 'Yellow' })
Write-Host ("  Failed       : {0}" -f $failCount) -ForegroundColor $(if ($failCount -gt 0) { 'Red' } else { 'Green' })
Write-Host ("  Total        : {0}" -f $level1Folders.Count)
Write-Host ""

if ($failCount -gt 0) {
    exit 1
}
