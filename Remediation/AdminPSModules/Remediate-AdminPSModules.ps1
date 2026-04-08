#Requires -Version 5.1
<#
.SYNOPSIS
    Intune Remediation - Remediation script
    Installs missing and updates outdated M365 admin PowerShell modules.

.DESCRIPTION
    For each module in the list, queries PSGallery for the latest version and either
    installs (if missing) or updates (if outdated) the module.
    Modules are installed in the AllUsers scope so they are available system-wide.

    Exit 0 = All modules successfully installed/updated
    Exit 1 = One or more modules failed - check Intune output log for details

.NOTES
    Author  : Jami Susijärvi, Oy Wolflake Consulting Ab
    GitHub  : https://github.com/Susijarvi
    Version : 1.0.0
    Created : 2026-04-05

    NOTE: Microsoft.Graph is a large meta-module. First-time installation
    can take 5-10 minutes. Intune remediation timeout is 30 minutes.

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
# If PS 7 is not installed, these modules are silently skipped.
#
# To enable a module: remove the leading #
# =============================================================================
$modulesPS7 = @(
    # 'ZeroTrustAssessment'  # Microsoft Zero Trust Assessment (requires PS 7.0+)
    # 'PnP.PowerShell'       # PnP PowerShell - M365 admin (SharePoint, Teams, etc.) (requires PS 7.2+)
)
# =============================================================================


# Log file - captures all output to disk for diagnosis.
# Intune's remediation output column is often empty even when the script writes to stdout.
$logFile = 'C:\Windows\Temp\Remediate-AdminPSModules.log'
New-Item -Path 'C:\Windows\Temp' -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
"$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') === Remediation started ===" |
    Out-File $logFile -Encoding utf8 -Force
function Write-Log {
    param([string]$Message)
    $line = "$(Get-Date -Format 'HH:mm:ss') $Message"
    Write-Output $line
    $line | Out-File $logFile -Append -Encoding utf8 -Force
}

# Catch any unhandled terminating error and exit 1 instead of crashing with an unexpected
# exit code that Intune reports as "Failed" rather than "Non-compliant"
trap {
    $line = "ERROR: Unhandled exception: $_"
    Write-Output $line
    try { $line | Out-File $logFile -Append -Encoding utf8 -Force } catch {}
    exit 1
}

# Force TLS 1.2 - required for PSGallery connections on older Windows versions
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Write-Log "PSModulePath: $($env:PSModulePath -replace ';', ' | ')"

# -AcceptLicense was added in PowerShellGet 1.6.0.
# PowerShellGet 1.0.0.1 (Windows built-in, present on Windows 365 cloud PCs) does not
# have this parameter - using it causes every Install-Module call to fail immediately.
$acceptLicenseSupported = (Get-Command Install-Module -ErrorAction SilentlyContinue).Parameters.ContainsKey('AcceptLicense')
Write-Log "AcceptLicense supported: $acceptLicenseSupported"
$installParams = @{
    Repository         = 'PSGallery'
    Scope              = 'AllUsers'
    Force              = $true
    AllowClobber       = $true
    SkipPublisherCheck = $true
    ErrorAction        = 'Stop'
}
if ($acceptLicenseSupported) { $installParams['AcceptLicense'] = $true }

$exitCode = 0
$pwsh7    = $null  # set below if PS7 modules are enabled and PS7 is found

# Ensure the NuGet package provider is installed and loaded in the current session.
# -ForceBootstrap suppresses the interactive ShouldContinue prompt that appears when
# NuGet is missing. Import-PackageProvider loads it into the current session so that
# Find-Module and Install-Module work without prompting in Intune SYSTEM context.
$nugetStep = 'init'
try {
    # Skip Get-PackageProvider check - it can trigger interactive ShouldContinue prompts
    # even with SilentlyContinue, because SilentlyContinue only suppresses terminating
    # errors, not ShouldContinue dialogs. Install-PackageProvider -Force is idempotent:
    # if NuGet is already installed at the required version it succeeds immediately.
    $nugetStep = 'Install-PackageProvider'
    Write-Log "Step 1: Installing/verifying NuGet package provider..."
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope AllUsers -Force -ForceBootstrap -ErrorAction Stop | Out-Null
    $nugetStep = 'Import-PackageProvider'
    Write-Log "Step 1a: Importing NuGet into session..."
    Import-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction Stop | Out-Null
    Write-Log "Step 1b: NuGet provider ready"
}
catch {
    Write-Log "ERROR: NuGet block failed at step '$nugetStep': $_"
    $exitCode = 1
}

# Trust PSGallery for the duration of this script to avoid install prompts
Write-Log "Step 2: Setting PSGallery as trusted..."
try {
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
    Write-Log "Step 2: PSGallery trusted"
}
catch {
    Write-Log "WARNING: Could not set PSGallery as trusted: $_"
}

# Locate PowerShell 7 if any PS7 modules are listed
if ($modulesPS7.Count -gt 0) {
    $pwshCmd = Get-Command pwsh -ErrorAction SilentlyContinue
    if ($pwshCmd) {
        $pwsh7 = $pwshCmd.Source
        Write-Log "PS7: Found PowerShell 7 at $pwsh7"
    } else {
        Write-Log "INFO: PS 7 not installed - PS7 modules will be skipped"
        Write-Log "INFO: Install PS 7 as an Intune Win32 app or manually: https://aka.ms/powershell"
    }
}

foreach ($moduleName in $modules) {
    try {
        # Fetch latest available version from PSGallery
        Write-Log "Querying PSGallery for: $moduleName"
        $latest = Find-Module -Name $moduleName -Repository PSGallery -ErrorAction Stop

        # Get the highest locally installed version
        $installed = Get-Module -ListAvailable -Name $moduleName |
            Sort-Object Version -Descending |
            Select-Object -First 1

        if (-not $installed) {
            Write-Log "Installing: $moduleName $($latest.Version)..."
            Install-Module -Name $moduleName @installParams
            $verify = Get-Module -ListAvailable -Name $moduleName |
                Sort-Object Version -Descending | Select-Object -First 1
            if ($verify) {
                Write-Log "OK: $moduleName $($latest.Version) installed - verified at $($verify.ModuleBase)"
            } else {
                Write-Log "WARNING: $moduleName claimed installed but NOT FOUND by Get-Module -ListAvailable"
                $exitCode = 1
            }
        }
        elseif ([version]$installed.Version -lt [version]$latest.Version) {
            Write-Log "Updating: $moduleName $($installed.Version) -> $($latest.Version)..."
            Update-Module -Name $moduleName `
                -Force `
                -ErrorAction Stop
            $verify = Get-Module -ListAvailable -Name $moduleName |
                Sort-Object Version -Descending | Select-Object -First 1
            if ($verify) {
                Write-Log "OK: $moduleName updated to $($latest.Version) - verified at $($verify.ModuleBase)"
            } else {
                Write-Log "WARNING: $moduleName claimed updated but NOT FOUND by Get-Module -ListAvailable"
                $exitCode = 1
            }
        }
        else {
            Write-Log "OK: $moduleName $($installed.Version) is up to date"
        }
    }
    catch {
        Write-Log "ERROR: $moduleName - $_"
        $exitCode = 1
    }
}

# Process PowerShell 7+ modules via pwsh subprocess
if ($modulesPS7.Count -gt 0 -and $pwsh7) {
    $ps7Base = "$env:ProgramFiles\PowerShell\Modules"
    foreach ($moduleName in $modulesPS7) {
        try {
            Write-Log "Querying PSGallery for PS7 module: $moduleName"
            $latest = Find-Module -Name $moduleName -Repository PSGallery -ErrorAction Stop

            $installedDir = Get-ChildItem "$ps7Base\$moduleName" -Directory -ErrorAction SilentlyContinue |
                Sort-Object { [version]$_.Name } -Descending | Select-Object -First 1

            if (-not $installedDir) {
                Write-Log "Installing (PS7): $moduleName $($latest.Version)..."
                & $pwsh7 -NonInteractive -NoProfile -Command "
                    Install-Module -Name '$moduleName' -Scope AllUsers -Force -AllowClobber -AcceptLicense -ErrorAction Stop
                " 2>&1 | ForEach-Object { Write-Log "  [pwsh] $_" }
                $verify = Get-ChildItem "$ps7Base\$moduleName" -Directory -ErrorAction SilentlyContinue |
                    Sort-Object { [version]$_.Name } -Descending | Select-Object -First 1
                if ($verify) {
                    Write-Log "OK: $moduleName $($latest.Version) installed (PS7)"
                } else {
                    Write-Log "WARNING: $moduleName not found after install (PS7)"
                    $exitCode = 1
                }
            }
            elseif ([version]$installedDir.Name -lt [version]$latest.Version) {
                Write-Log "Updating (PS7): $moduleName $($installedDir.Name) -> $($latest.Version)..."
                & $pwsh7 -NonInteractive -NoProfile -Command "
                    Update-Module -Name '$moduleName' -Force -ErrorAction Stop
                " 2>&1 | ForEach-Object { Write-Log "  [pwsh] $_" }
                Write-Log "OK: $moduleName updated to $($latest.Version) (PS7)"
            }
            else {
                Write-Log "OK: $moduleName $($installedDir.Name) is up to date (PS7)"
            }
        }
        catch {
            Write-Log "ERROR: $moduleName (PS7) - $_"
            $exitCode = 1
        }
    }
}

# Restore PSGallery to untrusted - leave the system in its original state
try {
    Set-PSRepository -Name PSGallery -InstallationPolicy Untrusted -ErrorAction SilentlyContinue
}
catch {}

exit $exitCode
