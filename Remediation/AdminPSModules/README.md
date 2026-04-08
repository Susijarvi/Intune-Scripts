# M365 Admin – PowerShell Modules (Install & Update)

Intune Proactive Remediation that ensures required M365 admin PowerShell modules
are installed and up to date on the device.

**Author:** Jami Susijärvi, Oy Wolflake Consulting Ab — [GitHub](https://github.com/Susijarvi)
**Version:** 1.0.0

---

## Scripts

| File | Purpose |
|------|---------|
| `Detect-AdminPSModules.ps1` | Checks if modules are installed and up to date. Exit 0 = compliant, Exit 1 = non-compliant. |
| `Remediate-AdminPSModules.ps1` | Installs missing modules and updates outdated ones. |
| `Manage-AdminPSModules.ps1` | Run manually in a console: shows color-coded status and optionally installs/updates modules interactively. Not used by Intune. |

---

## Intune Deployment Settings

> These settings are also documented inside each script's `.INTUNE` block.

**Intune > Devices > Scripts and remediations > Remediations > Create > Properties:**

| Field | Value |
|-------|-------|
| Package name | `M365 Admin - PowerShell Modules (Install & Update)` |
| Description | Ensures required M365 admin PowerShell modules are installed and up to date on the device. Checks ExchangeOnlineManagement, Microsoft.Graph, MicrosoftTeams, SharePoint Online, ORCA, ImportExcel, and PowerShellGet against the latest versions in PSGallery. Installs or updates any missing or outdated modules automatically. |
| Publisher | `Oy Wolflake Consulting Ab` |

**Scripts tab:**

| Setting | Value |
|---------|-------|
| Detection script | `Detect-AdminPSModules.ps1` |
| Remediation script | `Remediate-AdminPSModules.ps1` |
| Run script in 64-bit PowerShell host | **Yes** |
| Run this script using the logged-on credentials | **No** (runs as SYSTEM) |
| Enforce script signature check | **No** |

**Schedule tab:**
- Recommended: **Daily** or **Weekly** depending on environment

---

## Included Modules

### Active (managed by default)

| Module | Description |
|--------|-------------|
| `ExchangeOnlineManagement` | Exchange Online management |
| `Microsoft.Graph` | Microsoft Graph API — meta-module, installs all submodules |
| `MicrosoftTeams` | Microsoft Teams administration |
| `Microsoft.Online.SharePoint.PowerShell` | SharePoint Online management |
| `ORCA` | Office 365 Recommended Configuration Analyzer |
| `ImportExcel` | Excel report generation without Office installed |
| `PowerShellGet` | PowerShell module management |

### Optional PS 5.1 modules (commented out by default)

Uncomment in all three scripts to enable:

| Module | Description | Note |
|--------|-------------|------|
| `Az.Accounts` | Lightweight Azure authentication | Subset of the full Az module |
| `MicrosoftPowerBIMgmt` | Power BI administration | |
| `Microsoft.PowerApps.Administration.PowerShell` | Power Apps admin | |
| `WindowsAutopilotIntune` | Windows Autopilot management | |
| `IntuneBackupAndRestore` | Backup and restore Intune configurations | Community module |
| `PSWriteHTML` | Generate HTML reports from PowerShell | |

### Optional PS 7.0+ modules (requires PowerShell 7 pre-installed)

These modules require PowerShell 7 and cannot run in Intune's PS 5.1 SYSTEM context.
They are managed via a `pwsh` subprocess and processed automatically when `$modulesPS7` is non-empty.

| Module | Description | Minimum PS version |
|--------|-------------|-------------------|
| `ZeroTrustAssessment` | Microsoft Zero Trust Assessment | PS 7.0+ |
| `PnP.PowerShell` | M365 admin (SharePoint, Teams, etc.) | PS 7.2+ |

**To use PS 7 modules:**
1. Install PowerShell 7 on target devices first — recommended via Intune Win32 app
   (MSI installer: https://aka.ms/powershell)
2. Uncomment the desired modules in `$modulesPS7` in all three scripts (`Detect`, `Remediate`, `Manage`)

If PowerShell 7 is not installed on the device, PS7 modules are silently skipped
and the device remains compliant for the PS 5.1 modules.

---

## How to Enable or Disable a Module

Open all three scripts and comment/uncomment the module line. Example:

```powershell
# 'MicrosoftPowerBIMgmt'   # Power BI administration   <- disabled
'MicrosoftPowerBIMgmt'     # Power BI administration   <- enabled
```

All three scripts (`Detect`, `Remediate`, `Manage`) must have identical module lists to stay in sync.

---

## Logging

Both Intune scripts write a disk log during execution for diagnosis:

| Script | Log file |
|--------|----------|
| `Detect-AdminPSModules.ps1` | `C:\Windows\Temp\Detect-AdminPSModules.log` |
| `Remediate-AdminPSModules.ps1` | `C:\Windows\Temp\Remediate-AdminPSModules.log` |

The detection script also reads and re-emits the last remediation log in its output.
This is useful because Intune shows detection output reliably, but remediation output
is often empty in the Intune portal.

---

## Manual Management

Run `Manage-AdminPSModules.ps1` in an elevated PowerShell console:

```powershell
.\Manage-AdminPSModules.ps1
```

Output example:
```
 M365 Admin PowerShell Module Status
 Checked: 2026-04-07 12:00:00

Module                                     Installed      Latest         Status
------------------------------------------------------------------------------------------
ExchangeOnlineManagement                   3.9.2          3.9.2          OK
Microsoft.Graph                            2.36.1         2.36.1         OK
MicrosoftTeams                             7.6.0          7.7.0          OUTDATED
...

 2 module(s) need attention.
 Install/update now? [Y/N]:
```

If you choose Y, the script installs/updates the non-compliant modules and re-displays the status.

**Colors:**
- Green — installed and up to date
- Yellow — installed but a newer version is available
- Red — not installed

---

## Technical Notes

- Scripts run as **SYSTEM** via Intune — modules are installed to `AllUsers` scope
- `Microsoft.Graph` is a large meta-module — first-time installation can take **5–10 minutes**; Intune remediation timeout is 30 minutes
- NuGet package provider is installed automatically and silently before any PSGallery queries
- TLS 1.2 is enforced for PSGallery connections
- PS 7.0+ modules are managed via `pwsh` subprocess — Intune's PS 5.1 context cannot load them directly
- PowerShell 7 is **not** auto-installed by these scripts — deploy it separately as an Intune Win32 app

---

## Version History

| Version | Date | Description |
|---------|------|-------------|
| 1.0.0 | 2026-04-08 | Initial stable release |

---

## Disclaimer

These scripts are provided free of charge as a community tool, without warranty of any kind.
The author accepts no responsibility for any damage, data loss, or other consequences.
You run these scripts entirely at your own risk.
Always test in a non-production environment before deploying to production.
