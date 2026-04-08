#Requires -Version 5.1
<#
.SYNOPSIS
    Intune Remediation - Detection script
    Checks whether required M365 admin PowerShell modules are installed and up to date.

.DESCRIPTION
    Queries PSGallery for the latest version of each module and compares it against
    the locally installed version. If PSGallery is unreachable, the version check is
    skipped for that module (installation check still runs).

    Exit 0 = Compliant   - all modules installed and up to date
    Exit 1 = Non-compliant - one or more modules missing or outdated

.NOTES
    Author  : Jami Susijärvi, Oy Wolflake Consulting Ab
    GitHub  : https://github.com/Susijarvi
    Version : 1.1.0
    Created : 2026-04-05

.INTUNE
    Intune > Devices > Scripts and remediations > Remediations > Create > Properties:

      Package name        : M365 Admin - PowerShell Modules (Install & Update)
      Description         : Ensures required M365 admin PowerShell modules are installed and up
                            to date on the device. Checks ExchangeOnlineManagement, Microsoft.Graph,
                            MicrosoftTeams, SharePoint Online, PnP.PowerShell, ORCA,
                            ZeroTrustAssessment, ImportExcel, and PowerShellGet against the latest
                            versions in PSGallery. Installs or updates any missing or outdated
                            modules automatically.
      Publisher           : Oy Wolflake Consulting Ab

    Scripts tab:
      Detection script    : Detect-AdminPSModules.ps1
      Remediation script  : Remediate-AdminPSModules.ps1

      Run script in 64-bit PowerShell host              : Yes
      Run this script using the logged-on credentials   : No  (runs as SYSTEM)
      Enforce script signature check                    : No

    Schedule tab:
      Recommended         : Daily or Weekly depending on environment

.DISCLAIMER
    This script is provided free of charge as a community tool, without warranty of
    any kind - express or implied. The author accepts no responsibility for any damage,
    data loss, misconfiguration, or other consequences arising from the use of this
    script. You run this script entirely at your own risk. Always test in a non-production
    environment before deploying to production.

.CHANGELOG
    1.0.0 - 2026-04-08 - Initial stable release. Manages M365 admin PowerShell modules via
                          Intune Proactive Remediation — silently detects, remediates, and
                          interactively manages module installation and updates.
    1.1.0 - 2026-04-08 - Version bump to stay in sync with Manage (scope selection added).
#>

# =============================================================================
# MODULE LIST
# Add or remove modules here. Comment out (#) any module you don't need.
# To add a new module, append a new line following the same format.
#
# OPTIONAL modules are listed below the active ones and commented out by default.
# Remove the leading # to enable them.
# =============================================================================
$modules = @(

    # --- Active modules ---
    'ExchangeOnlineManagement'               # Exchange Online management
    'Microsoft.Graph'                        # Microsoft Graph API (meta-module, installs all submodules)
    'MicrosoftTeams'                         # Microsoft Teams administration
    'Microsoft.Online.SharePoint.PowerShell' # SharePoint Online management
    'ORCA'                                   # Office 365 Recommended Configuration Analyzer
    'ImportExcel'                            # Excel report generation without Office installed
    'PowerShellGet'                          # PowerShell module management (keep up to date)

    # --- Optional modules - uncomment to enable ---
    # 'Az.Accounts'                          # Lightweight Azure authentication (subset of the full Az module)
    # 'MicrosoftPowerBIMgmt'                 # Power BI administration
    # 'Microsoft.PowerApps.Administration.PowerShell' # Power Apps admin
    # 'WindowsAutopilotIntune'               # Windows Autopilot management via Intune
    # 'IntuneBackupAndRestore'               # Backup and restore Intune configurations (community module)
    # 'PSWriteHTML'                          # Generate HTML reports from PowerShell output

)
# =============================================================================

# =============================================================================
# MODULE LIST - PowerShell 7.0+ modules (optional)
#
# These modules require PowerShell 7 and are managed via a pwsh subprocess.
# PowerShell 7 must be pre-installed on the device before enabling this.
# Recommended: deploy PS 7 as an Intune Win32 app (MSI from https://aka.ms/powershell)
# If PS 7 is not installed, these modules are silently skipped (device stays compliant).
#
# To enable a module: remove the leading #
# =============================================================================
$modulesPS7 = @(
    # 'ZeroTrustAssessment'  # Microsoft Zero Trust Assessment (requires PS 7.0+)
    # 'PnP.PowerShell'       # PnP PowerShell - M365 admin (SharePoint, Teams, etc.) (requires PS 7.2+)
)
# =============================================================================


# Log file - captures all output to disk for diagnosis.
$logFile = 'C:\Windows\Temp\Detect-AdminPSModules.log'
New-Item -Path 'C:\Windows\Temp' -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
"$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') === Detection started ===" |
    Out-File $logFile -Encoding utf8 -Force
function Write-Log {
    param([string]$Message)
    $line = "$(Get-Date -Format 'HH:mm:ss') $Message"
    Write-Output $line
    $line | Out-File $logFile -Append -Encoding utf8 -Force
}

# Catch any unhandled terminating error and exit 1 (non-compliant) instead of crashing
# with an unexpected exit code that Intune reports as "Failed" rather than "Non-compliant"
trap {
    $line = "ERROR: Unhandled exception: $_"
    Write-Output $line
    try { $line | Out-File $logFile -Append -Encoding utf8 -Force } catch {}
    exit 1
}

# Force TLS 1.2 - required for PSGallery connections on older Windows versions
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Emit the last remediation log in detection output so it is visible in Intune.
# Intune shows detection output but often hides remediation output - this bridges the gap.
$remLogFile = 'C:\Windows\Temp\Remediate-AdminPSModules.log'
if (Test-Path $remLogFile) {
    Write-Output "--- Last remediation log ---"
    Get-Content $remLogFile | ForEach-Object { Write-Output $_ }
    Write-Output "--- End remediation log ---"
}

$nonCompliant = [System.Collections.Generic.List[string]]::new()

# Ensure NuGet provider is installed and loaded in the current session.
# -ForceBootstrap suppresses the interactive ShouldContinue prompt that appears when
# NuGet is missing - this is what caused the "Failed" status in Intune SYSTEM context.
# Import-PackageProvider loads the provider into the current session after installation,
# otherwise Find-Module can still prompt even after a successful install.
$nugetReady = $true
try {
    # Skip Get-PackageProvider check - it can trigger interactive ShouldContinue prompts
    # even with SilentlyContinue, because SilentlyContinue only suppresses terminating
    # errors, not ShouldContinue dialogs. Install-PackageProvider -Force is idempotent:
    # if NuGet is already installed at the required version it succeeds immediately.
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope AllUsers -Force -ForceBootstrap -ErrorAction Stop | Out-Null
    Import-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction Stop | Out-Null
}
catch {
    Write-Log "WARNING: NuGet provider installation failed - version checks will be skipped: $_"
    $nugetReady = $false
}

foreach ($moduleName in $modules) {
    # Get the highest installed version of the module
    $installed = Get-Module -ListAvailable -Name $moduleName |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if (-not $installed) {
        $nonCompliant.Add("$moduleName (not installed)")
        continue
    }

    # Compare installed version against latest version on PSGallery
    if (-not $nugetReady) {
        # NuGet unavailable - cannot query PSGallery, skip version check
        continue
    }

    try {
        $latest = Find-Module -Name $moduleName -Repository PSGallery -ErrorAction Stop
        if ([version]$installed.Version -lt [version]$latest.Version) {
            $nonCompliant.Add("$moduleName (installed: $($installed.Version), latest: $($latest.Version))")
        }
    }
    catch {
        # PSGallery unreachable - skip version check, installation already confirmed above
        Write-Log "WARNING: PSGallery query failed for $moduleName - skipping version check"
    }
}

# Check PowerShell 7+ modules
if ($modulesPS7.Count -gt 0) {
    $pwshCmd = Get-Command pwsh -ErrorAction SilentlyContinue
    if (-not $pwshCmd) {
        # PS7 not installed - skip silently, device stays compliant.
        # Install PS 7 separately (e.g. Intune Win32 app) to manage these modules.
        Write-Log "INFO: PS 7 not installed - PS7 modules skipped"
    } else {
        $ps7Base = "$env:ProgramFiles\PowerShell\Modules"
        foreach ($moduleName in $modulesPS7) {
            $installedDir = Get-ChildItem "$ps7Base\$moduleName" -Directory -ErrorAction SilentlyContinue |
                Sort-Object { [version]$_.Name } -Descending | Select-Object -First 1

            if (-not $installedDir) {
                $nonCompliant.Add("$moduleName (not installed) [PS7]")
                continue
            }

            if (-not $nugetReady) { continue }

            try {
                $latest = Find-Module -Name $moduleName -Repository PSGallery -ErrorAction Stop
                if ([version]$installedDir.Name -lt [version]$latest.Version) {
                    $nonCompliant.Add("$moduleName (installed: $($installedDir.Name), latest: $($latest.Version)) [PS7]")
                }
            } catch {
                Write-Log "WARNING: PSGallery query failed for $moduleName (PS7) - skipping version check"
            }
        }
    }
}

if ($nonCompliant.Count -gt 0) {
    Write-Log "NON-COMPLIANT: $($nonCompliant -join ' | ')"
    exit 1
}

Write-Log "COMPLIANT: All modules are installed and up to date"
exit 0
