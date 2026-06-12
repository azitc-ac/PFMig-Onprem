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
$level1Folders = @(Get-PublicFolder -Recurse |
    Where-Object { $_.ParentPath -eq "\" } |
    Select-Object Name, ParentPath)

if ($level1Folders.Count -eq 0) {
    Write-Host "No Level-1 folders found."
    exit 0
}

Write-Host "Found $($level1Folders.Count) Level-1 folder(s)`n" -ForegroundColor Green

# Pre-check: determine which folders are mail-enabled (including subfolders)
Write-Host "Checking mail-enabled status (including subfolders)..." -ForegroundColor Yellow
$folderStatus = @()
$folderHierarchy = @()

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
        IsSubfolder = $false
    }

    $folderHierarchy += [pscustomobject]@{
        Path = $sourcePath
        MailEnabled = $isMailEnabled
        Level = 0
        Children = @()
    }

    # Check subfolders
    try {
        $subfolders = @(Get-PublicFolder -Identity $sourcePath -Recurse -ErrorAction Stop |
            Where-Object { $_.ParentPath -like "$sourcePath*" -and $_.ParentPath -ne $sourcePath })

        foreach ($subfolder in $subfolders) {
            $subPath = $subfolder.ParentPath + "\" + $subfolder.Name
            $subIsMailEnabled = $false

            try {
                $subMailpf = Get-MailPublicFolder -Identity $subPath -ErrorAction SilentlyContinue
                $subIsMailEnabled = $true
            } catch {
                # Not mail-enabled
            }

            $folderStatus += [pscustomobject]@{
                Name = $subfolder.Name
                Path = $subPath
                MailEnabled = $subIsMailEnabled
                IsSubfolder = $true
            }

            $lastHierarchy = $folderHierarchy | Where-Object { $_.Path -eq $sourcePath } | Select-Object -Last 1
            if ($lastHierarchy) {
                $lastHierarchy.Children += [pscustomobject]@{
                    Path = $subPath
                    MailEnabled = $subIsMailEnabled
                }
            }
        }
    } catch {
        # Subfolder enumeration error
    }
}

Write-Host ""
Write-Host "Folder Hierarchy (mail-enabled status):" -ForegroundColor Cyan
foreach ($item in $folderHierarchy) {
    $mailIcon = if ($item.MailEnabled) { "[✓]" } else { "[ ]" }
    Write-Host ("  {0,-50} {1}" -f $item.Path, $mailIcon) -ForegroundColor $(if ($item.MailEnabled) { 'Green' } else { 'Gray' })

    foreach ($child in $item.Children) {
        $childMailIcon = if ($child.MailEnabled) { "[✓]" } else { "[ ]" }
        Write-Host ("    ├─ {0,-46} {1}" -f $child.Path, $childMailIcon) -ForegroundColor $(if ($child.MailEnabled) { 'Yellow' } else { 'Gray' })
    }
}
Write-Host ""

# Summary of mail-enabled folders for cutover
$mailEnabledCount = @($folderStatus | Where-Object { $_.MailEnabled }).Count
if ($mailEnabledCount -gt 0) {
    Write-Host "Mail-enabled folders detected: $mailEnabledCount" -ForegroundColor Green
    Write-Host "These will be considered for mailflow cutover if -CutoverMailflow is used.`n" -ForegroundColor Green
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
