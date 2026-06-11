<#
    Orchestration script for Public Folder migration (on-prem, WIA)
    Refactor: 2026-01
    Requires: PSmig-Functions.Refactored.ps1 in same directory
    Notes:
      - PS 5.1 compatible, supports -WhatIf / -Confirm
      - OGV for PF/OU selection retained
      - TLS validation bypass enabled (like original)
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [string]$PublicFolderPath,
    [string]$OUforMBX,
    [string]$TargetMailboxName,
    [switch]$DoNotCopyItems,
    [switch]$CopyOnly,
    [switch]$CutoverMailflow,
    [ValidateSet('FullMailbox','FolderACL')][string]$PermissionMode = 'FullMailbox',
    [string]$Url,
    [string]$AutodiscoverUrl,
    [string]$AdditionalAdminSam
)

# Load functions
. "$PSScriptRoot\PFmig-Functions.ps1"

cd $PSScriptRoot
$start = Get-Date
$log   = Join-Path $PSScriptRoot ("PFMig_{0:yyyyMMdd_HHmmss}.log" -f $start)
Start-Transcript -Path $log -Append | Out-Null

# Version-bewusstes CSV-Encoding: PS 5.1 schreibt mit 'UTF8' bereits BOM,
# PS 6+/7 braucht 'utf8BOM' fuer dasselbe Ergebnis.
$csvEncoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }

try {
    Write-Host ("`nPublic Folder Migration  -  {0:yyyy-MM-dd HH:mm}" -f $start) -ForegroundColor Green

    # Prepare Exchange cmdlets (snapin if needed)
    $snapin = Get-PSSnapin -Registered Microsoft.Exchange.Management.PowerShell.E* -ErrorAction SilentlyContinue
    if ($snapin -eq $null) { Write-Warning "Snapin not available" }
    else {
        $cmd = Get-Command "Get-Mailbox" -ErrorAction SilentlyContinue
        if ($cmd -eq $null) {
            Write-Verbose "Loading Exchange Commands"
            Add-PSSnapin $snapin -ErrorAction SilentlyContinue
        } else {
            Write-Verbose "Exchange Snapin is already loaded"
        }
    }

    Set-ADServerSettings -ViewEntireForest:$true

    $adminSam = "$env:USERDOMAIN\$env:USERNAME"
    $adminUser = Get-User -Identity $adminSam
    $adminUPN  = $adminUser.UserPrincipalName
    Write-Verbose "Running as [$adminSam]"

# Get PF path (OGV retained, robust)
if (-not $PublicFolderPath) {
    # Liste vollständig materialisieren (verhindert Timing-/Lazy-Probleme)
    $pfList = @(
        Get-PublicFolder -Recurse |
        Where-Object { $_.ParentPath } |
        Select-Object Name, ParentPath
    )

    # Auswahl EXPLIZIT übernehmen
    $selection = $pfList | Out-GridView -PassThru -Title "Select the source PF structure"

    if (-not $selection) { Write-Host "Abgebrochen."; return }

    # Wenn mehrere gewählt wurden → ersten verwenden
    if ($selection -is [System.Array]) {
        $selection = $selection | Select-Object -First 1
    }

    # Pfad bestimmen
    if ($selection.Name -eq "IPM_SUBTREE") {
        $PublicFolderPath = "\"
    } else {
        if ($selection.ParentPath -ne "\") {
            $PublicFolderPath = ($selection.ParentPath + "\" + $selection.Name)
        } else {
            $PublicFolderPath = "\" + $selection.Name
        }
    }

    Write-Verbose ("Source PF path: {0}" -f $PublicFolderPath)
}

    # Derive/create target mailbox
    if (-not $TargetMailboxName) {
        $last = ($PublicFolderPath.Split("\") | Select-Object -Last 1)
        if ([string]::IsNullOrWhiteSpace($last)) { $last = "ROOT" }
        $TargetMailboxName = ("PF_{0}" -f $last) `
            -ireplace 'ä','ae' -ireplace 'ö','oe' -ireplace 'ü','ue' -ireplace 'ß','ss' `
            -ireplace ' ', '' -ireplace '\\','_'
    }

    Write-Verbose "Checking/creating target mailbox..."
    $mailboxJustCreated = $false
    $mbx = Get-Mailbox -Identity $TargetMailboxName -ErrorAction SilentlyContinue
    if (-not $mbx) {
        $mailboxJustCreated = $true
        if (-not $OUforMBX) {
            # Kein OU-DN per Parameter -> gespeicherten Pfad anbieten bzw. OGV-Auswahl
            # (siehe Get-TargetOU). Auswahl wird in dieser .config-Datei persistiert.
            $ouConfig = Join-Path $PSScriptRoot 'PFMig.OU.config'
            $ouDn = Get-TargetOU -ConfigPath $ouConfig
            if (-not $ouDn) { Write-Host "Cancelled."; return }
        }
        else {
            $ouDn = $OUforMBX
        }
        if ($PSCmdlet.ShouldProcess(("Mailbox '{0}'" -f $TargetMailboxName),("Create in OU '{0}'" -f $ouDn))) {
            #New-Mailbox -Name $TargetMailboxName -OrganizationalUnit $ouDn -Shared -Database "Postfachspeicher01_MGH" | Out-Null
            New-Mailbox -Name $TargetMailboxName -OrganizationalUnit $ouDn -Shared | Out-Null
        }

        # wait until resolvable
        do {
            Write-Host "Mailbox creation started, checking progress in 10 seconds..."
            Start-Sleep 10
            $mbx = Get-Mailbox -Identity $TargetMailboxName -ErrorAction SilentlyContinue
        } while (-not $mbx)

        # admin FullAccess (no automapping)
        if ($PSCmdlet.ShouldProcess(("Mailbox '{0}'" -f $TargetMailboxName),"Grant FullAccess to current admin")) {
            Add-MailboxPermission -Identity $TargetMailboxName -User $adminSam -AccessRights FullAccess -AutoMapping:$false -WarningAction SilentlyContinue | Out-Null
        }
    } else {
        Write-Verbose "Using existing mailbox [$TargetMailboxName]"
    }

    # Robustly determine SMTP: WindowsEmailAddress may be empty on newly created mailboxes → fallback to PrimarySmtpAddress
    function Resolve-MailboxSmtp {
        param($Mailbox)
        $smtp = $null
        if ($Mailbox.WindowsEmailAddress -and $Mailbox.WindowsEmailAddress.Address) {
            $smtp = $Mailbox.WindowsEmailAddress.Address
        }
        if ([string]::IsNullOrWhiteSpace($smtp) -and $Mailbox.PrimarySmtpAddress) {
            $smtp = $Mailbox.PrimarySmtpAddress.ToString()
        }
        return $smtp
    }

    $targetAlias = $mbx.Alias

    # Wait for valid target SMTP (address policy can take a few seconds to apply)
    $targetSmtp = Resolve-MailboxSmtp -Mailbox $mbx
    $waited = 0
    while ([string]::IsNullOrWhiteSpace($targetSmtp) -and $waited -lt 60) {
        Write-Verbose "Waiting for target mailbox SMTP address..."
        Start-Sleep 5; $waited += 5
        $mbx = Get-Mailbox -Identity $TargetMailboxName -ErrorAction SilentlyContinue
        $targetSmtp = Resolve-MailboxSmtp -Mailbox $mbx
    }
    if ([string]::IsNullOrWhiteSpace($targetSmtp)) {
        throw ("Could not determine SMTP address for target mailbox '{0}'." -f $TargetMailboxName)
    }

    $adminMbx = Get-Mailbox -Identity $adminSam -ErrorAction SilentlyContinue
    if (-not $adminMbx) { throw ("Admin mailbox '{0}' not found (EWS connection requires a mailbox)." -f $adminSam) }
    $adminMbxSmtp = Resolve-MailboxSmtp -Mailbox $adminMbx
    if ([string]::IsNullOrWhiteSpace($adminMbxSmtp)) {
        throw ("Could not determine SMTP address for admin mailbox '{0}'." -f $adminSam)
    }

    # PF folders to process
    $folders = Get-PublicFolder $PublicFolderPath -Recurse | Where-Object { $_.ParentPath }
    Write-Verbose (($folders | Format-Table | Out-String))

    Write-Host ""
    Write-Host ("Source : {0}" -f $PublicFolderPath)
    Write-Host ("Target : {0}  ({1})" -f $TargetMailboxName, $targetSmtp)

    # Two separate EWS instances
    # $service      → PF access (with PF routing headers set)
    # $targetService → target mailbox access with impersonation
    $service       = Connect-Exchange -MailboxName $adminMbxSmtp -Url $Url
    $targetService = Connect-Exchange -MailboxName $adminMbxSmtp -Url $Url

    # Set impersonation on target mailbox (requires ApplicationImpersonation role)
    $targetService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId(
        [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,
        $targetSmtp
    )
    Write-Verbose ("EWS impersonation set for [{0}]" -f $targetSmtp)

    # Derive Autodiscover URL once from working EWS host (if not explicitly provided via -AutodiscoverUrl)
    if (-not $AutodiscoverUrl) {
        $AutodiscoverUrl = Resolve-AutodiscoverUrl -Service $service
        if ($AutodiscoverUrl) {
            Write-Verbose ("Autodiscover URL derived: {0}" -f $AutodiscoverUrl)
        } else {
            Write-Warning "No Autodiscover URL derivable - PF routing headers will be attempted via SCP lookup."
        }
    } else {
        Write-Verbose ("Autodiscover-URL (Parameter): {0}" -f $AutodiscoverUrl)
    }

    # PREFLIGHT 1: Check ApplicationImpersonation role (warn, don't fail – effective user resolution can be unreliable)
    try {
        $impAssign = Get-ManagementRoleAssignment -Role ApplicationImpersonation -GetEffectiveUsers -ErrorAction Stop |
                     Where-Object { $_.EffectiveUserName -eq $adminUser.Name -or $_.EffectiveUserName -eq $adminSam }
        if (-not $impAssign) {
            Write-Warning (("Could not confirm ApplicationImpersonation role for [{0}]. " -f $adminSam) +
                "If target access fails: New-ManagementRoleAssignment -Role ApplicationImpersonation -User '$adminSam' (RBAC cache ~15 min).")
        } else {
            Write-Verbose ("ApplicationImpersonation role confirmed for [{0}]." -f $adminSam)
        }
    } catch {
        Write-Verbose ("RBAC preflight skipped: {0}" -f $_.Exception.Message)
    }

    # PREFLIGHT 2: Actively test EWS access to target (Bind + Retry)
    # Waits for Store provisioning of newly created mailbox AND validates impersonation
    # On permanent 401/impersonation error: fail fast with clear message instead of cryptic downstream errors
    Write-Verbose "Testing EWS access to target mailbox (impersonation)..."
    # Newly created mailboxes take longer for Store to provision → larger time window (~5 min) vs standard 2 min
    $accessAttempts = if ($mailboxJustCreated) { 30 } else { 12 }
    Test-EwsMailboxAccess -Service $targetService -MailboxSmtp $targetSmtp -Context 'Target mailbox' -MaxAttempts $accessAttempts | Out-Null

    Write-Host ""
    Write-Host ("Migrating {0} folders..." -f $folders.Count)
    $folderIndex = 0

    # Collect pf addresses during loop
    [System.Collections.ArrayList]$pfAddresses = @()
    $primarySmtp = $null

    # Endabrechnung: pro verarbeitetem Ordner ein Ergebnisobjekt + neu erstellte Elternordner
    $migrationResults = New-Object System.Collections.ArrayList
    $createdParents   = New-Object System.Collections.Generic.List[string]

    foreach ($f in $folders) {
        $folderIndex++
        $fullPath = (($f.ParentPath + "\" + $f.Name).Replace("\\","\"))  # normalized
        $class    = $f.FolderClass

        Write-Verbose ("[{0}/{1}] Processing [{2}] - FolderClass [{3}]" -f $folderIndex, $folders.Count, $fullPath, $class)

        # Redetermine PF routing headers per iteration (prevents stale headers if folders in different PF mailboxes)
        foreach ($hdr in @('X-AnchorMailbox','X-PublicFolderMailbox')) {
            # [void]: .Remove() returns bool, suppress from output stream
            if ($service.HttpHeaders.ContainsKey($hdr)) { [void]$service.HttpHeaders.Remove($hdr) }
        }


        # Create parent folder chain and set traversal permissions
        $parentPath = $f.ParentPath

        if ($parentPath -and $parentPath -ne '\') {
            # Split path safely. Explicit [string[]] typing prevents PS 5.1 scalar coercion
            # when Where-Object returns single element. Example: '\Folder1\Folder1.2' → @('Folder1','Folder1.2')
            [string[]]$segments = $parentPath.Trim('\').Split('\') | Where-Object { $_ -ne '' } | ForEach-Object { [string]$_ }

            $parentAgg      = $null     # kumulierter Pfad: '\Folder1', dann '\Folder1\Folder1.2'
            $parentForCreate = '\'      # Start-Eltern ist der Root

            foreach ($seg in $segments) {
                # 1) Create parent folder (idempotent) via targetService (no PF routing)
                # Capture return: reports if NEW created (no pipeline leak)
                $cfParent = Create-Folder -MailboxName $targetSmtp -Service $targetService -NewFolderName $seg -ParentFolder $parentForCreate

                # 2) Ensure traversal (Default = Reviewer) on current parent
                $pathForCmd = ('{0}:{1}' -f $targetAlias, $parentForCreate)
                if ($PSCmdlet.ShouldProcess(("Folder '{0}'" -f $pathForCmd), "Ensure Default Reviewer (traversal)")) {
                    try {
                        Set-MailboxFolderPermission -Identity $pathForCmd -User "Default" -AccessRights Reviewer -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
                    } catch {
                        # bereits gesetzt / nicht anwendbar -> ignorieren
                    }
                }

                # 3) Update aggregated parent path: new parent is just-created folder
                if ([string]::IsNullOrEmpty($parentAgg)) {
                    $parentAgg = "\" + $seg
                } else {
                    $parentAgg = $parentAgg + "\" + $seg
                }
                $parentForCreate = $parentAgg

                # Track newly created parent folders for final summary
                if ($cfParent -and $cfParent.Created -and -not $createdParents.Contains($parentAgg)) {
                    $createdParents.Add($parentAgg)
                }
            }
        }
        else {
            # Top-level case: no parent chain, only traversal on root (harmless if already set)
            $pathForCmd = ('{0}:{1}' -f $targetAlias, '\')
            if ($PSCmdlet.ShouldProcess(("Folder '{0}'" -f $pathForCmd), "Ensure Default Reviewer (root traversal)")) {
                try {
                    Set-MailboxFolderPermission -Identity $pathForCmd -User "Default" -AccessRights Reviewer -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
                } catch { }
            }
        }


        # Copy content (or just ensure folder exists) – Ergebnis fuer Endabrechnung erfassen
        $pfiResult = Get-PublicFolderItems -Service $service `
            -TargetService $targetService `
            -AdminMailboxSmtp $adminMbxSmtp `
            -PublicFolderPath $fullPath `
            -TargetMailboxSmtp $targetSmtp `
            -ParentPath $f.ParentPath `
            -FolderName $f.Name `
            -FolderClass $class `
            -AutodiscoverUrl $AutodiscoverUrl `
            -DoNotCopyItems:$DoNotCopyItems
        # Defensive: nur echte Ergebnisobjekte aufnehmen (falls eine EWS-Methode wider
        # Erwarten doch einmal etwas in den Output-Stream leakt).
        $pfiResult = @($pfiResult) |
            Where-Object { $_ -and ($_.PSObject.Properties.Name -contains 'ItemsCopied') } |
            Select-Object -Last 1
        if ($pfiResult) {
            [void]$migrationResults.Add($pfiResult)

            # Eine ruhige Statuszeile pro Ordner. Farbe nur bei Problem.
            $problem = ($pfiResult.ItemsFailed -gt 0) -or ($pfiResult.Status -notin @('OK','NurOrdner'))
            $mark    = if ($problem) { '!' } else { 'ok' }
            $line = ("  [{0}/{1}] {2,-40} kopiert {3}, uebersprungen {4}, Fehler {5}  {6}" -f `
                        $folderIndex, $folders.Count, $fullPath, `
                        $pfiResult.ItemsCopied, $pfiResult.ItemsSkipped, $pfiResult.ItemsFailed, $mark)
            if ($problem) { Write-Host $line -ForegroundColor Yellow } else { Write-Host $line }
        }

        # Mail-enabled PF?
        Write-Verbose "Collecting mail addresses..."  # Already in English
        try {
            $mailpf = Get-MailPublicFolder -Identity $fullPath -ErrorAction Stop
        } catch {
            $mailpf = $null
            Write-Verbose ("Get-MailPublicFolder: {0}" -f $_.Exception.Message)
        }

        if ($mailpf) {
            $asStrings = [string]::Join(';', ($mailpf.EmailAddresses.AddressString))
            Write-Verbose ("Folder [{0}] has these mail addresses: [{1}]" -f $fullPath,$asStrings)
            foreach ($smtp in $mailpf.EmailAddresses) { [void]$pfAddresses.Add($smtp.AddressString) }
            $primarySmtp = $mailpf.WindowsEmailAddress.Address

            # Send-As perms
            Write-Verbose "Reading and copying SendAs permissions..."
            $perms = $mailpf | Get-ADPermission
            $sendAs = $perms | Where-Object { $_.ExtendedRights -like "Send-As" }
            foreach ($p in $sendAs) {
                $user = $p.User
                Write-Verbose ("Copying Send-As permission for [{0}] to [{1}]" -f $user, $targetAlias)
                $mbxDn = (Get-Mailbox -Identity $TargetMailboxName).DistinguishedName
                if ($PSCmdlet.ShouldProcess(("Mailbox '{0}'" -f $TargetMailboxName),("Grant Send-As to '{0}'" -f $user))) {
                    Add-ADPermission -Identity $mbxDn -ExtendedRights 'Send-As' -User $user -WarningAction SilentlyContinue | Out-Null
                }
            }
        }

        # Folder permissions
        Write-Verbose "Migrating / Copying folder permissions..."
        $perms = Get-PublicFolderClientPermission -Identity $fullPath
        foreach ($perm in $perms) {
            $user = $perm.User
            $rights = $perm.AccessRights

            # Default/Anonymous werden auf Ordner gesetzt (nicht auf Mailbox)
            if ($user.DisplayName -in @('None','Default','Standard','Anonymous','Anonym')) {
                $pathForCmd = ('{0}:{1}' -f $targetAlias, $fullPath)
                if ($PSCmdlet.ShouldProcess(("Folder '{0}'" -f $pathForCmd),("Set permission for '{0}'" -f $user))) {
                    try {
                        Set-MailboxFolderPermission -Identity $pathForCmd -User $user -AccessRights $rights -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
                    } catch {
                        Write-Warning $_.Exception.Message
                    }
                }
            } else {
                if ($PermissionMode -eq 'FullMailbox') {
                    # Einfach: FullAccess aufs Postfach
                    $who = $null
                    try {
                        $who = $user.ADRecipient.WindowsEmailAddress.Address
                    } catch {
                        $who = $user
                    }
                    if ($who) {
                        if ($PSCmdlet.ShouldProcess(("Mailbox '{0}'" -f $TargetMailboxName),("Grant FullAccess to '{0}'" -f $who))) {
                            if ($CopyOnly) {
                                Add-MailboxPermission -Identity $TargetMailboxName -AccessRights FullAccess -User $who -WhatIf
                            } else {
                                Add-MailboxPermission -Identity $TargetMailboxName -AccessRights FullAccess -User $who -WarningAction SilentlyContinue | Out-Null
                            }
                        }
                    }
                } else {
                    # Präzise: ACL auf Zielordner
                    $pathForCmd = ('{0}:{1}' -f $targetAlias, $fullPath)
                    if ($PSCmdlet.ShouldProcess(("Folder '{0}'" -f $pathForCmd),("Add folder permission for '{0}'" -f $user))) {
                        try {
                            Add-MailboxFolderPermission -Identity $pathForCmd -User $user -AccessRights $rights -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
                        } catch {
                            Write-Warning $_.Exception.Message
                        }
                    }
                }
            }
        }

        # Additional admin as Owner (optional)
        if ($AdditionalAdminSam) {
            $pathForCmd = ('{0}:{1}' -f $targetAlias, $fullPath)
            if ($PSCmdlet.ShouldProcess(("Folder '{0}'" -f $pathForCmd),("Grant Owner to '{0}'" -f $AdditionalAdminSam))) {
                try {
                    Add-MailboxFolderPermission -Identity $pathForCmd -User $AdditionalAdminSam -AccessRights Owner -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
                } catch {
                    Write-Warning $_.Exception.Message
                }
            }
        }

        # Mailflow Cutover (disable mail-enabled PF)
        if ($mailpf) {
            if ($CutoverMailflow) {
                Write-Verbose ("Disabling mail-enabled PF [{0}]..." -f $fullPath)
                if ($PSCmdlet.ShouldProcess(("Mail Public Folder '{0}'" -f $fullPath),"Disable")) {
                    Disable-MailPublicFolder -Identity $mailpf -Confirm:$false
                }
                Start-Sleep 10
            } else {
                Write-Verbose "WHATIF: would disable mail PF [$fullPath]"
            }
        }
    } # foreach folder

    # Assign PF addresses to target mailbox after cutover
    if ($CutoverMailflow) {
        Write-Verbose "Assigning mail addresses of the disabled PFs to the target mailbox..."
        if ($pfAddresses.Count -gt 0) {
            $uniq = $pfAddresses | Sort-Object -Unique
            $normalized = foreach ($a in $uniq) {
                if ($a -match '^(?i)smtp:(.*)$') {
                    $addr = $Matches[1]
                    if ($primarySmtp -and ($addr -ieq $primarySmtp)) { "SMTP:$addr" } else { "smtp:$addr" }
                } else { $a }  # X500:, SIP:, etc.
            }

            if ($PSCmdlet.ShouldProcess(("Mailbox '{0}'" -f $TargetMailboxName),"Set proxy addresses / primary SMTP")) {
                Set-Mailbox -Identity $TargetMailboxName -EmailAddresses $normalized -EmailAddressPolicyEnabled:$false
                if ($primarySmtp) {
                    Set-Mailbox -Identity $TargetMailboxName -WindowsEmailAddress $primarySmtp
                }
            }
        }
    } else {
        if ($pfAddresses.Count -gt 0) {
            $sim = $pfAddresses -join '; '
            Write-Verbose ("WHATIF: would set EmailAddresses on [{0}] to: {1}" -f $TargetMailboxName,$sim)
        }
    }

    # ============================ ENDABRECHNUNG ============================
    # Neu erstellte Ordner (Eltern + Blatt) sammeln
    $allCreated = New-Object System.Collections.Generic.List[string]
    foreach ($p in $createdParents) { if (-not $allCreated.Contains($p)) { $allCreated.Add($p) } }
    foreach ($r in $migrationResults) {
        if ($r.FolderCreated -and -not $allCreated.Contains($r.TargetPath)) { $allCreated.Add($r.TargetPath) }
    }

    # Summen
    $sumCopied = 0; $sumSkipped = 0; $sumFailed = 0
    foreach ($r in $migrationResults) {
        if ($null -eq $r) { continue }
        $sumCopied  += [int]$r.ItemsCopied
        $sumSkipped += [int]$r.ItemsSkipped
        $sumFailed  += [int]$r.ItemsFailed
    }

    # Detail-Tabelle nur bei -Verbose
    if ($migrationResults.Count -gt 0) {
        Write-Verbose (($migrationResults |
            Select-Object @{N='Quelle';E={$_.SourcePath}}, @{N='Ziel';E={$_.TargetPath}},
                          @{N='Neu';E={$_.FolderCreated}}, @{N='Kopiert';E={$_.ItemsCopied}},
                          @{N='Skip';E={$_.ItemsSkipped}}, @{N='Fehler';E={$_.ItemsFailed}},
                          @{N='Status';E={$_.Status}} |
            Format-Table -AutoSize | Out-String))
    }

    Write-Host "`nZusammenfassung"
    Write-Host ("  Ordner verarbeitet : {0}" -f $migrationResults.Count)
    if ($allCreated.Count -gt 0) {
        Write-Host ("  Neue Ordner        : {0}" -f $allCreated.Count)
        foreach ($c in $allCreated) { Write-Host ("      + {0}" -f $c) }
    } else {
        Write-Host  "  Neue Ordner        : 0 (alle bereits vorhanden)"
    }
    Write-Host ("  Items kopiert      : {0}  (uebersprungen {1}, Fehler {2})" -f $sumCopied, $sumSkipped, $sumFailed) `
        -ForegroundColor $(if ($sumFailed -gt 0) { 'Yellow' } else { 'Green' })

    if ($DoNotCopyItems) {
        Write-Host "  Hinweis            : -DoNotCopyItems aktiv (nur Ordnerstruktur)."
    } elseif ($sumCopied -eq 0 -and $sumSkipped -eq 0) {
        Write-Host "  ACHTUNG            : Es wurde NICHTS kopiert (0 Items). Quell-Zugriff/Berechtigung pruefen, oder Quellordner leer." -ForegroundColor Red
    } elseif ($sumCopied -eq 0 -and $sumSkipped -gt 0) {
        Write-Host "  Hinweis            : 0 neu kopiert - alle Items laut Kopier-Log bereits vorhanden."
    }

    $failedFolders = @($migrationResults | Where-Object { $_.Status -notin @('OK','NurOrdner','TeilweiseFehlgeschlagen') })
    if ($failedFolders.Count -gt 0) {
        Write-Host ("  Ordner mit Fehlern : {0}" -f $failedFolders.Count) -ForegroundColor Red
        foreach ($ff in $failedFolders) {
            Write-Host ("      ! {0}  [{1}] {2}" -f $ff.SourcePath, $ff.Status, $ff.Error) -ForegroundColor Red
        }
    }

    # 4) Report als CSV neben das Transcript schreiben
    $reportCsv = Join-Path $PSScriptRoot ("PFMig_{0:yyyyMMdd_HHmmss}_report.csv" -f $start)
    $reportRows = New-Object System.Collections.ArrayList
    # neu erstellte Elternordner als eigene Zeilen aufnehmen (sonst nicht in den Ergebnissen)
    foreach ($p in $createdParents) {
        [void]$reportRows.Add([pscustomobject]@{
            SourcePath = '(Elternordner)'; TargetPath = $p; FolderCreated = $true
            ItemsCopied = 0; ItemsSkipped = 0; ItemsFailed = 0
            Status = 'ElternordnerErstellt'; Error = $null
        })
    }
    foreach ($r in $migrationResults) { [void]$reportRows.Add($r) }

    if ($reportRows.Count -gt 0) {
        try {
            $reportRows |
                Select-Object SourcePath, TargetPath, FolderCreated, ItemsCopied, ItemsSkipped, ItemsFailed, Status, Error |
                Export-Csv -Path $reportCsv -NoTypeInformation -Encoding $csvEncoding -ErrorAction Stop
            Write-Host ("  Report             : {0}" -f $reportCsv)
        } catch {
            Write-Warning ("Konnte Report '{0}' nicht schreiben: {1}" -f $reportCsv, $_.Exception.Message)
        }
    }
    # =======================================================================

    $end = Get-Date
    $dur = $end - $start
    Write-Host ("  Dauer              : {0:00}:{1:00}:{2:00}" -f $dur.Hours,$dur.Minutes,$dur.Seconds)
    Write-Host "Fertig." -ForegroundColor Green
}
finally {
    Stop-Transcript | Out-Null
}
