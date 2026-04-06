# M365 Admin – PowerShell Modules (Install & Update)

Intune Proactive Remediation that ensures required M365 admin PowerShell modules
are installed and up to date on the device.

**Author:** Jami Susijärvi, Oy Wolflake Consulting Ab
**Version:** 0.8 (Detect) / 0.6 (Remediate)
**Status:** In development — not yet at 1.0 release

---

## Scripts

| File | Purpose |
|------|---------|
| `Detect-AdminPSModules.ps1` | Checks if modules are installed and up to date. Exit 0 = compliant, Exit 1 = non-compliant. |
| `Remediate-AdminPSModules.ps1` | Installs missing modules and updates outdated ones. |
| `Get-AdminPSModuleStatus.ps1` | Run manually in a console to get a color-coded status overview. Not used by Intune. |

---

## Intune Deployment Settings

> These settings are also documented inside each script's `.INTUNE` block.

**Intune > Devices > Remediations > Create > Properties:**

| Field | Value |
|-------|-------|
| Package name | `M365 Admin - PowerShell Modules (Install & Update)` |
| Description | Ensures required M365 admin PowerShell modules are installed and up to date on the device. Checks ExchangeOnlineManagement, Microsoft.Graph, MicrosoftTeams, SharePoint Online, ORCA, ZeroTrustAssessment, ImportExcel, and PowerShellGet against the latest versions in PSGallery. Installs or updates any missing or outdated modules automatically. |
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

### Active (installed by default)

| Module | Description |
|--------|-------------|
| `ExchangeOnlineManagement` | Exchange Online management |
| `Microsoft.Graph` | Microsoft Graph API — meta-module, installs all submodules |
| `MicrosoftTeams` | Microsoft Teams administration |
| `Microsoft.Online.SharePoint.PowerShell` | SharePoint Online management |
| `ORCA` | Office 365 Recommended Configuration Analyzer |
| `ZeroTrustAssessment` | Microsoft Zero Trust Assessment tool |
| `ImportExcel` | Excel report generation without Office installed |
| `PowerShellGet` | PowerShell module management |

### Optional (commented out by default)

Uncomment in the module list to enable:

| Module | Description | Note |
|--------|-------------|------|
| `AzureAD` | Azure AD management | DEPRECATED — but still widely used |
| `MSOnline` | Legacy MSOL module | DEPRECATED — still needed in some environments |
| `Az.Accounts` | Lightweight Azure authentication | Subset of the full Az module |
| `MicrosoftPowerBIMgmt` | Power BI administration | |
| `Microsoft.PowerApps.Administration.PowerShell` | Power Apps admin | |
| `WindowsAutopilotIntune` | Windows Autopilot management | |
| `IntuneBackupAndRestore` | Backup and restore Intune configurations | Community module |
| `PSWriteHTML` | Generate HTML reports from PowerShell | |

To enable an optional module, open all three scripts and remove the `#` from the module line:
```powershell
# 'MicrosoftPowerBIMgmt'   # Power BI administration
```
becomes:
```powershell
'MicrosoftPowerBIMgmt'     # Power BI administration
```

---

## Manual Status Check

Run `Get-AdminPSModuleStatus.ps1` locally to see a color-coded overview before or after deployment:

```powershell
.\Get-AdminPSModuleStatus.ps1
```

Output example:
```
 M365 Admin PowerShell Module Status
 Checked: 2026-04-06 12:00:00

Module                                     Installed      Latest         Status
------------------------------------------------------------------------------------------
ExchangeOnlineManagement                   3.9.2          3.9.2          OK
Microsoft.Graph                            2.36.1         2.36.1         OK
ZeroTrustAssessment                        -              ?              NOT INSTALLED
...
```

**Colors:**
- 🟢 Green — installed and up to date
- 🟡 Yellow — installed but a newer version is available
- 🔴 Red — not installed

---

## Technical Notes

- Scripts run as **SYSTEM** via Intune — modules are installed to `AllUsers` scope
- `Microsoft.Graph` is a large meta-module — first-time installation can take **5–10 minutes**
- `PnP.PowerShell` is excluded — it requires PowerShell 7.2+, incompatible with Intune's PS 5.1 context
- NuGet package provider is installed automatically and silently before any PSGallery queries
- TLS 1.2 is enforced for PSGallery connections

---

## Version History

| Version | Date | Description |
|---------|------|-------------|
| 0.1 | 2026-04-05 | Initial version |
| 0.2 | 2026-04-05 | Added optional modules section |
| 0.3 | 2026-04-05 | Fixed interactive NuGet prompt in status script |
| 0.4 | 2026-04-06 | Removed PnP.PowerShell (PS 7.2+ requirement) |
| 0.5 | 2026-04-06 | Added -ForceBootstrap, TLS 1.2, global trap, -AcceptLicense |
| 0.6 | 2026-04-06 | Removed Get-PackageProvider pre-check (prompt fix for Windows 365) |
| 0.7 | 2026-04-06 | Version sync |
| 0.8 | 2026-04-06 | Detect: removed Get-PackageProvider check — idempotent install |

> Version 1.0 will be tagged when the scripts are confirmed stable across all target environments.

---

## Disclaimer

These scripts are provided free of charge as a community tool, without warranty of any kind.
The author accepts no responsibility for any damage, data loss, or other consequences.
You run these scripts entirely at your own risk.
Always test in a non-production environment before deploying to production.
