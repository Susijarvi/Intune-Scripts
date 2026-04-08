#Requires -Version 5.1
<#
.SYNOPSIS
    M365 Admin PowerShell Module Manager
    Shows module status and optionally installs or updates modules interactively.

.DESCRIPTION
    Run this script manually in an elevated PowerShell console to check the status
    of all required M365 admin PowerShell modules and optionally install or update them.

    Phase 1 - Status: Shows installed vs. latest PSGallery version for each module,
              color-coded (Green = OK, Yellow = outdated, Red = not installed).

    Phase 2 - Install/update: If any modules need attention, offers to install or
              update them. Handles both PS 5.1 and PS 7.0+ modules.

    Phase 3 - Re-check: Displays updated status after install/update.

.NOTES
    Author  : Jami Susijärvi, Oy Wolflake Consulting Ab
    GitHub  : https://github.com/Susijarvi
    Version : 1.1.0
    Created : 2026-04-07

    Run as administrator to install modules in AllUsers scope.

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
    1.1.0 - 2026-04-08 - Added scope selection: choose CurrentUser or AllUsers at runtime.
#>

# =============================================================================
# MODULE LIST - keep in sync with Detect-AdminPSModules.ps1
#
# OPTIONAL modules are listed below the active ones and commented out by default.
# Remove the leading # to enable them.
# =============================================================================
$modules = @(

    # --- Active modules ---
    'ExchangeOnlineManagement'               # Exchange Online management
    'Microsoft.Graph'                        # Microsoft Graph API (meta-module)
    'MicrosoftTeams'                         # Microsoft Teams administration
    'Microsoft.Online.SharePoint.PowerShell' # SharePoint Online management
    'ORCA'                                   # Office 365 Recommended Configuration Analyzer
    'ImportExcel'                            # Excel report generation without Office installed
    'PowerShellGet'                          # PowerShell module management

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
# These modules require PowerShell 7. To enable a module: remove the leading #
# =============================================================================
$modulesPS7 = @(
    # 'ZeroTrustAssessment'  # Microsoft Zero Trust Assessment (requires PS 7.0+)
    # 'PnP.PowerShell'       # PnP PowerShell - M365 admin (SharePoint, Teams, etc.) (requires PS 7.2+)
)
# =============================================================================

# Force TLS 1.2 - required for PSGallery connections
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# =============================================================================
# SCOPE SELECTION
# CurrentUser: installs to your profile, roams with OneDrive, no admin required
# AllUsers:    installs system-wide, requires admin
# Note: NuGet provider always installs to AllUsers (CurrentUser not supported)
# =============================================================================
Write-Host ""
Write-Host " Install scope:" -ForegroundColor White
Write-Host "  [1] CurrentUser  - installs to your profile, roams with OneDrive, no admin required (default)"
Write-Host "  [2] AllUsers     - installs system-wide, requires admin"
$scopeAnswer = Read-Host " Select [1/2]"
$installScope = if ($scopeAnswer -eq '2') { 'AllUsers' } else { 'CurrentUser' }
Write-Host (" Using scope: {0}" -f $installScope) -ForegroundColor Cyan

# =============================================================================
# NUGET PROVIDER CHECK
# Required to query PSGallery for latest versions.
# =============================================================================
$nugetReady = $true
$nuget = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
if (-not $nuget -or $nuget.Version -lt '2.8.5.201') {
    Write-Host ""
    Write-Host " NuGet provider is required to query PSGallery for latest versions." -ForegroundColor DarkYellow
    $answer = Read-Host " Install NuGet provider now? [Y/N]"
    if ($answer -match '^[Yy]') {
        try {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope AllUsers -Force -ErrorAction Stop | Out-Null
            Import-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction Stop | Out-Null
            Write-Host " NuGet provider installed." -ForegroundColor Green
        }
        catch {
            Write-Host " WARNING: NuGet provider installation failed - version checks will be skipped." -ForegroundColor DarkYellow
            $nugetReady = $false
        }
    }
    else {
        Write-Host " Skipping NuGet install - version checks will be omitted." -ForegroundColor DarkYellow
        $nugetReady = $false
    }
}

# =============================================================================
# DISPLAY HELPERS
# =============================================================================
$colModule    = 42
$colInstalled = 14
$colLatest    = 14
$colStatus    = 20

function Write-Row {
    param(
        [string]$Module,
        [string]$Installed,
        [string]$Latest,
        [string]$Status,
        [System.ConsoleColor]$Color
    )
    $line = "{0,-$colModule} {1,-$colInstalled} {2,-$colLatest}" -f $Module, $Installed, $Latest
    Write-Host $line -NoNewline
    Write-Host $Status -ForegroundColor $Color
}

function Write-TableHeader {
    Write-Host ""
    $header = "{0,-$colModule} {1,-$colInstalled} {2,-$colLatest} {3,-$colStatus}" -f "Module", "Installed", "Latest", "Status"
    Write-Host $header -ForegroundColor White
    Write-Host ("-" * ($colModule + $colInstalled + $colLatest + $colStatus + 3)) -ForegroundColor DarkGray
}

function Write-TableFooter {
    Write-Host ("-" * ($colModule + $colInstalled + $colLatest + $colStatus + 3)) -ForegroundColor DarkGray
}

# =============================================================================
# STATUS CHECK FUNCTION
# Returns a hashtable of modules needing action: @{ ModuleName = 'install'|'update' }
# =============================================================================
function Get-ModuleStatus {
    param([switch]$PS7)

    $needsAction = [System.Collections.Generic.Dictionary[string,string]]::new()
    $countOk = 0; $countOutdated = 0; $countMissing = 0

    $list    = if ($PS7) { $modulesPS7 } else { $modules }
    $ps7Base = "$env:ProgramFiles\PowerShell\Modules"

    foreach ($moduleName in $list) {

        if ($PS7) {
            $installedDir = Get-ChildItem "$ps7Base\$moduleName" -Directory -ErrorAction SilentlyContinue |
                Sort-Object { [version]$_.Name } -Descending | Select-Object -First 1
            $installedVersion = if ($installedDir) { $installedDir.Name } else { $null }
        } else {
            $installed = Get-Module -ListAvailable -Name $moduleName |
                Sort-Object Version -Descending | Select-Object -First 1
            $installedVersion = if ($installed) { $installed.Version.ToString() } else { $null }
        }

        if (-not $installedVersion) {
            Write-Row -Module $moduleName -Installed "-" -Latest "?" -Status "NOT INSTALLED" -Color Red
            $needsAction[$moduleName] = 'install'
            $countMissing++
            continue
        }

        if (-not $nugetReady) {
            Write-Row -Module $moduleName -Installed $installedVersion -Latest "-" -Status "VERSION SKIPPED" -Color DarkYellow
            continue
        }

        try {
            $latest = Find-Module -Name $moduleName -Repository PSGallery -ErrorAction Stop
            if ([version]$installedVersion -lt [version]$latest.Version) {
                Write-Row -Module $moduleName -Installed $installedVersion -Latest $latest.Version -Status "OUTDATED" -Color Yellow
                $needsAction[$moduleName] = 'update'
                $countOutdated++
            } else {
                Write-Row -Module $moduleName -Installed $installedVersion -Latest $latest.Version -Status "OK" -Color Green
                $countOk++
            }
        }
        catch {
            Write-Row -Module $moduleName -Installed $installedVersion -Latest "?" -Status "GALLERY UNREACHABLE" -Color DarkYellow
        }
    }

    Write-TableFooter
    return @{ NeedsAction = $needsAction; Ok = $countOk; Outdated = $countOutdated; Missing = $countMissing }
}

# =============================================================================
# PHASE 1 - STATUS DISPLAY
# =============================================================================
Write-Host ""
Write-Host " M365 Admin PowerShell Module Status" -ForegroundColor Cyan
Write-Host (" Checked: {0}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss")) -ForegroundColor Cyan

Write-TableHeader
$ps5Result  = Get-ModuleStatus
$ps5Needs   = $ps5Result.NeedsAction

Write-Host ""
Write-Host " Summary (PS 5.1 modules):" -ForegroundColor White
Write-Host ("  OK           : {0}" -f $ps5Result.Ok)       -ForegroundColor Green
Write-Host ("  Outdated     : {0}" -f $ps5Result.Outdated)  -ForegroundColor Yellow
Write-Host ("  Not installed: {0}" -f $ps5Result.Missing)   -ForegroundColor Red

$ps7Needs   = [System.Collections.Generic.Dictionary[string,string]]::new()
$pwsh7      = $null

if ($modulesPS7.Count -gt 0) {
    Write-Host ""
    Write-Host " PowerShell 7 Modules" -ForegroundColor Cyan
    $pwshCmd = Get-Command pwsh -ErrorAction SilentlyContinue
    if (-not $pwshCmd) {
        Write-Host " PowerShell 7 is not installed - PS7 modules cannot be checked." -ForegroundColor DarkYellow
        Write-Host " Install PS 7 from: https://aka.ms/powershell" -ForegroundColor DarkYellow
    } else {
        $pwsh7 = $pwshCmd.Source
        Write-Host (" PowerShell 7: {0}" -f $pwsh7) -ForegroundColor DarkGray
        Write-TableHeader
        $ps7Result = Get-ModuleStatus -PS7
        $ps7Needs  = $ps7Result.NeedsAction
        Write-Host ""
        Write-Host " Summary (PS 7 modules):" -ForegroundColor White
        Write-Host ("  OK           : {0}" -f $ps7Result.Ok)       -ForegroundColor Green
        Write-Host ("  Outdated     : {0}" -f $ps7Result.Outdated)  -ForegroundColor Yellow
        Write-Host ("  Not installed: {0}" -f $ps7Result.Missing)   -ForegroundColor Red
    }
}

# =============================================================================
# PHASE 2 - INSTALL / UPDATE PROMPT
# =============================================================================
$totalNeeds = $ps5Needs.Count + $ps7Needs.Count
if ($totalNeeds -eq 0) {
    Write-Host ""
    Write-Host " All modules are up to date. Nothing to do." -ForegroundColor Green
    Write-Host ""
    exit 0
}

Write-Host ""
Write-Host (" {0} module(s) need attention." -f $totalNeeds) -ForegroundColor Yellow
$answer = Read-Host " Install/update now? [Y/N]"
if ($answer -notmatch '^[Yy]') {
    Write-Host " Skipped." -ForegroundColor DarkYellow
    Write-Host ""
    exit 0
}

# Setup PSGallery trust + AcceptLicense check
try {
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
} catch {}

$acceptLicenseSupported = (Get-Command Install-Module -ErrorAction SilentlyContinue).Parameters.ContainsKey('AcceptLicense')
$installParams = @{
    Scope        = $installScope
    Force        = $true
    AllowClobber = $true
    ErrorAction  = 'Stop'
}
if ($acceptLicenseSupported) { $installParams['AcceptLicense'] = $true }

Write-Host ""

# --- PS 5.1 modules ---
foreach ($moduleName in $ps5Needs.Keys) {
    $action = $ps5Needs[$moduleName]
    try {
        if ($action -eq 'install') {
            Write-Host " Installing: $moduleName..." -ForegroundColor Cyan -NoNewline
            Install-Module -Name $moduleName @installParams
            Write-Host " done." -ForegroundColor Green
        } else {
            Write-Host " Updating:   $moduleName..." -ForegroundColor Cyan -NoNewline
            Update-Module -Name $moduleName -Force -ErrorAction Stop
            Write-Host " done." -ForegroundColor Green
        }
    }
    catch {
        Write-Host " FAILED: $_" -ForegroundColor Red
    }
}

# --- PS 7 modules ---
if ($ps7Needs.Count -gt 0 -and $pwsh7) {
    $ps7Base = "$env:ProgramFiles\PowerShell\Modules"
    foreach ($moduleName in $ps7Needs.Keys) {
        $action = $ps7Needs[$moduleName]
        try {
            if ($action -eq 'install') {
                Write-Host " Installing (PS7): $moduleName..." -ForegroundColor Cyan -NoNewline
                & $pwsh7 -NonInteractive -NoProfile -Command "
                    Install-Module -Name '$moduleName' -Scope $installScope -Force -AllowClobber -AcceptLicense -ErrorAction Stop
                " 2>&1 | Where-Object { $_ } | ForEach-Object { Write-Host "  [pwsh] $_" }
                Write-Host " done." -ForegroundColor Green
            } else {
                Write-Host " Updating (PS7):   $moduleName..." -ForegroundColor Cyan -NoNewline
                & $pwsh7 -NonInteractive -NoProfile -Command "
                    Update-Module -Name '$moduleName' -Scope $installScope -Force -ErrorAction Stop
                " 2>&1 | Where-Object { $_ } | ForEach-Object { Write-Host "  [pwsh] $_" }
                Write-Host " done." -ForegroundColor Green
            }
        }
        catch {
            Write-Host " FAILED: $_" -ForegroundColor Red
        }
    }
}

# Restore PSGallery
try { Set-PSRepository -Name PSGallery -InstallationPolicy Untrusted -ErrorAction SilentlyContinue } catch {}

# =============================================================================
# PHASE 3 - RE-CHECK STATUS
# =============================================================================
Write-Host ""
Write-Host " Re-checking status..." -ForegroundColor Cyan

Write-Host ""
Write-Host " M365 Admin PowerShell Module Status" -ForegroundColor Cyan
Write-Host (" Checked: {0}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss")) -ForegroundColor Cyan

Write-TableHeader
Get-ModuleStatus | Out-Null

if ($modulesPS7.Count -gt 0 -and $pwsh7) {
    Write-Host ""
    Write-Host " PowerShell 7 Modules" -ForegroundColor Cyan
    Write-Host (" PowerShell 7: {0}" -f $pwsh7) -ForegroundColor DarkGray
    Write-TableHeader
    Get-ModuleStatus -PS7 | Out-Null
}

Write-Host ""
