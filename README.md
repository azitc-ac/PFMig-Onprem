# Public Folder to Shared Mailbox Migration (On-Premises Exchange)

Robust PowerShell toolset for migrating Exchange public folder structure and content to shared mailboxes on-premises using EWS Managed API.

## Overview

This project provides two complementary scripts for public folder migration:

- **`Migrate-PublicFolders.ps1`** – Orchestrator script. Orchestrates the end-to-end migration: folder hierarchy creation, item copying, permission migration, and mailflow cutover.
- **`PFmig-Functions.ps1`** – Function library. Provides helpers for EWS connectivity, folder management, item copying, and error handling.

**Key Features:**
- ✅ Interactive public folder and OU selection (with caching)
- ✅ Automatic shared mailbox creation with SMTP provisioning
- ✅ Per-item error handling (one item failure doesn't abort migration)
- ✅ Oversized/corrupt item filtering via `MaxItemSize` (skipped, not failed)
- ✅ Incremental item copying via CSV deduplication log
- ✅ Folder hierarchy and traversal permission management
- ✅ Mail-enabled PF address migration and mailflow cutover
- ✅ Active mailbox access testing with retry logic
- ✅ Robust error classification (auth vs. transient vs. store unavailability)
- ✅ Full transcript logging and CSV report generation
- ✅ PowerShell 5.1 compatible

## Requirements

### Software
- **PowerShell 5.1** or later (on Windows)
- **EWS Managed API 2.2** (Microsoft.Exchange.WebServices.dll)
  - Separate component – **not** installed with Exchange Server by default
  - The old Microsoft download (`details.aspx?id=42951`) has been **retired**; obtain the DLL via **NuGet** (see [Installing the EWS Managed API](#installing-the-ews-managed-api))
- **Windows Integrated Authentication** with Exchange

### Permissions
- **ApplicationImpersonation RBAC role** assigned to running user account
  ```powershell
  New-ManagementRoleAssignment -Role ApplicationImpersonation -User <domain\username>
  ```
  (RBAC cache can take ~15 minutes to propagate)
- Read access to source public folders
- Mailbox creation rights in target OU

### Environment
- Exchange 2013 SP1 or later (on-premises)
- Running on Exchange server or with Management Tools installed

## Installation

```powershell
git clone https://github.com/azitc-ac/PFMig-Onprem.git
cd PFMig-Onprem
```

Then install the EWS Managed API (see below).

## Installing the EWS Managed API

The Microsoft download page for the EWS Managed API 2.2 MSI has been retired. The assembly is now distributed via **NuGet** (`Microsoft.Exchange.WebServices`, v2.2.0).

The script's loader (`Load-EWSManagedAPI`) searches for `Microsoft.Exchange.WebServices.dll` in this order:
1. Explicit path passed to the function (`-DllPath`)
2. Environment variable `EWS_MANAGED_API_DLL`
3. **Next to the scripts** (script folder)
4. A **`lib`** subfolder next to the scripts
5. Legacy MSI install (registry `HKLM\SOFTWARE\Microsoft\Exchange\Web Services`)

So the simplest setup is: get the DLL via NuGet and drop it into the repo's `lib` folder.

### Option A – `nuget.exe` (no admin rights needed)
```powershell
# Download nuget.exe if you don't have it: https://www.nuget.org/downloads
.\nuget.exe install Microsoft.Exchange.WebServices -Version 2.2.0 -OutputDirectory .\packages

# Copy the DLL into a lib folder next to the scripts
New-Item -ItemType Directory -Force -Path .\lib | Out-Null
Copy-Item .\packages\Microsoft.Exchange.WebServices.2.2.0\lib\40\Microsoft.Exchange.WebServices.dll .\lib\
```

### Option B – PowerShell `Install-Package` (PackageManagement)
```powershell
# Register the NuGet source once if needed:
# Register-PackageSource -Name nuget.org -Location https://api.nuget.org/v3/index.json -ProviderName NuGet

Install-Package Microsoft.Exchange.WebServices -RequiredVersion 2.2.0 -Source nuget.org -Scope CurrentUser -Force

# Locate and copy the DLL into the lib folder
$pkg = (Get-Package Microsoft.Exchange.WebServices).Source | Split-Path
New-Item -ItemType Directory -Force -Path .\lib | Out-Null
Copy-Item (Join-Path $pkg 'lib\40\Microsoft.Exchange.WebServices.dll') .\lib\
```

### Option C – point to an existing DLL
If the DLL already exists somewhere on the machine (e.g. an old MSI install or another app):
```powershell
$env:EWS_MANAGED_API_DLL = 'C:\Path\To\Microsoft.Exchange.WebServices.dll'
```

> **Important – unblock the DLL.** A DLL downloaded from the internet (NuGet) is marked as blocked by Windows and .NET will refuse to load it. Clear the mark of the web:
> ```powershell
> Unblock-File .\lib\Microsoft.Exchange.WebServices.dll
> ```

### Verify it loads
```powershell
. .\PFmig-Functions.ps1
Load-EWSManagedAPI -Verbose
# -> "EWS Managed API loaded from: ...\lib\Microsoft.Exchange.WebServices.dll"
```

## Usage

### Basic (Interactive)
```powershell
.\Migrate-PublicFolders.ps1
```
Prompts for folder selection, mailbox name, and target OU.

### With Mailflow Cutover
```powershell
.\Migrate-PublicFolders.ps1 -CutoverMailflow
```
Disables mail-enabled PF and assigns addresses to target mailbox.

### Incremental (Skip Already-Copied Items)
```powershell
# Run multiple times; CSV dedup log skips already-copied items
.\Migrate-PublicFolders.ps1
```

### Specific Folder
```powershell
.\Migrate-PublicFolders.ps1 `
  -PublicFolderPath '\Folder1\Folder1.1' `
  -TargetMailboxName 'SharedMBX_Folder1.1' `
  -OUforMBX 'OU=SharedMailboxes,DC=example,DC=com'
```

## Parameters

| Name | Type | Default | Description |
|------|------|---------|-------------|
| PublicFolderPath | string | (OGV select) | Source PF path |
| OUforMBX | string | (OGV select) | OU for shared mailbox (cached) |
| TargetMailboxName | string | (auto-derived) | Mailbox name |
| DoNotCopyItems | switch | false | Folder structure only |
| CopyOnly | switch | false | Skip permissions |
| CutoverMailflow | switch | false | Disable PF, assign addresses |
| PermissionMode | string | FullMailbox | FullMailbox or FolderACL |
| Url | string | https://COMPUTERNAME/EWS/Exchange.asmx | Explicit EWS endpoint (defaults to local server) |
| AutodiscoverUrl | string | (auto-derive) | Explicit Autodiscover URL |
| AdditionalAdminSam | string | (none) | Additional Owner |
| MaxItemSize | long | 10MB | Items larger than this are skipped (logged as oversized) and do not abort the run |

## Handling Oversized / Corrupt Items

Large or corrupt items can cause EWS `Copy()` to fail with **"Fehler beim Verschiebungs- oder Kopiervorgang"** ("Error during the move or copy operation"). To prevent these from aborting the run, the script pre-checks each item's size:

- Items larger than `-MaxItemSize` (default **10 MB**) are **skipped** and counted as *oversized* (not as failures).
- Skipped items are recorded in the dedup log with status `Oversized`, so subsequent runs don't re-attempt them.
- Items that still fail to copy (e.g. genuinely corrupt) are logged with status `Failed` and listed in `PFMig_<timestamp>_failed.csv` for review.

**Raise or lower the threshold** as needed:
```powershell
# Allow items up to 20 MB
.\Migrate-PublicFolders.ps1 -PublicFolderPath '\Wartungskalender' -MaxItemSize 20MB
```

After the run, review which items were skipped:
```powershell
Import-Csv .\PFMig_<timestamp>_oversized.csv | Format-Table Subject, Size
Import-Csv .\PFMig_<timestamp>_failed.csv   | Format-Table Subject, ErrorMessage
```

## Output

- **Console:** Concise status with per-folder item counts (copied / skipped / oversized / failed)
- **Transcript:** `PFMig_<timestamp>.log` – full details
- **Report:** `PFMig_<timestamp>_report.csv` – status per folder
- **Oversized Report:** `PFMig_<timestamp>_oversized.csv` – items skipped for exceeding `MaxItemSize`
- **Failed Report:** `PFMig_<timestamp>_failed.csv` – items that could not be copied, with error message
- **Dedup Log:** `<mailbox-smtp>.csv` – prevents re-copying items (incl. oversized/failed)

## Troubleshooting

### EWS connection failed
```powershell
.\Migrate-PublicFolders.ps1 -Url 'https://localhost/EWS/Exchange.asmx'
```

### ApplicationImpersonation denied (401)
```powershell
New-ManagementRoleAssignment -Role ApplicationImpersonation -User "$env:USERDOMAIN\$env:USERNAME"
# Wait ~15 minutes
```

### Quick Check: Verify ApplicationImpersonation Works
Before running a full migration, verify that ApplicationImpersonation is functioning:

```powershell
# Load EWS Managed API (adjust path to where you placed the DLL, e.g. .\lib\)
Add-Type -Path '.\lib\Microsoft.Exchange.WebServices.dll'

# Connect to EWS with default credentials (Windows Integrated Auth)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
$service.UseDefaultCredentials = $true
$service.Url = [System.Uri]'https://localhost/EWS/Exchange.asmx'

# Test impersonation: try to access target mailbox folder
try {
    $rootId = New-Object Microsoft.Exchange.WebServices.Data.FolderId(
        [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, 
        'target-mailbox@domain.com'  # Replace with target mailbox SMTP
    )
    [void][Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $rootId)
    Write-Host "✅ ApplicationImpersonation is working!" -ForegroundColor Green
} catch {
    Write-Host "❌ ApplicationImpersonation failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
```

If this fails with `401 Unauthorized`, the RBAC role assignment hasn't propagated yet or the user doesn't have the role. Wait 15 minutes and retry.

### Mailbox Store temporarily unavailable
- Wait 2–5 minutes
- Re-run script (uses existing mailbox, retries access)
- Check service health:
  ```powershell
  Get-Service MSExchangeIS | Format-Table Name, Status
  Get-MailboxDatabase -Status | Format-Table Name, Mounted
  ```

### No items copied
1. Check source permissions
2. Verify source folder not empty: `Get-PublicFolder <path> -ResultSize Unlimited | Get-PublicFolderItem`
3. Check dedup log: `<target-mailbox-smtp>.csv`

### "Fehler beim Verschiebungs- oder Kopiervorgang" (Error during move/copy)
Usually caused by items that are too large or corrupt for EWS `Copy()`.
- The script skips items above `-MaxItemSize` (default 10 MB) automatically – see [Handling Oversized / Corrupt Items](#handling-oversized--corrupt-items).
- If many items still fail, lower the threshold or inspect `PFMig_<timestamp>_failed.csv`.
  ```powershell
  .\Migrate-PublicFolders.ps1 -PublicFolderPath '\YourFolder' -MaxItemSize 5MB
  ```

## Architecture

**Two separate scripts:**
- `PFmig-Functions.ps1` – Pure helpers (EWS, folders, items, error handling)
- `Migrate-PublicFolders.ps1` – Orchestration only (preflight, mailbox, permissions, reports)

**Error handling:**
- **Auth failures** → Fail fast with diagnostic hints
- **Transient errors** → Retry with backoff
- **Per-item failures** → Log warning, skip item, continue

**Two EWS connections:**
- `$service` – PF access (with routing headers)
- `$targetService` – Mailbox access (with impersonation)

## Known Limitations

1. **Domain Controller:** Store provisioning delays common (unsupported config)
2. **TLS bypass:** Trust-all policy enabled for lab/test (not production safe)
3. **PS 5.1 edge cases:** Path segment scalar coercion fixed with explicit casting
4. **No hierarchy-only:** Parent folders must be migrated first

## License

[MIT](LICENSE)

## Authors

- **Glen Scales** – Original EWS migration scripts
- **Alexander Zarenko** – Refactor 2026-01: robustness, English translation, documentation

---

**Version:** 2.0 (Refactored, English, Production-Ready)  
**Updated:** 2026-06-16
