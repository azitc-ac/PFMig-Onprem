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
  - [Download](https://www.microsoft.com/en-us/download/details.aspx?id=42951)
  - Or obtain via Exchange on-premises installation
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
git clone https://github.com/azitc-ac/PF2SharedMBXOnprem.git
cd PF2SharedMBXOnprem
```

Verify EWS Managed API:
```powershell
Get-ItemProperty "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services" |
  Get-ChildItem | Sort-Object Name -Descending | Select-Object -First 1
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
| Url | string | (Autodiscover) | Explicit EWS endpoint |
| AutodiscoverUrl | string | (auto-derive) | Explicit Autodiscover URL |
| AdditionalAdminSam | string | (none) | Additional Owner |

## Output

- **Console:** Concise status with per-folder item counts
- **Transcript:** `PFMig_<timestamp>.log` – full details
- **Report:** `PFMig_<timestamp>_report.csv` – status per folder
- **Dedup Log:** `<mailbox-smtp>.csv` – prevents re-copying items

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
**Updated:** 2026-06-11
