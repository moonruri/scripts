<#
.SYNOPSIS
    Check-WindowsUpdate.ps1
    Checks if Windows Update is managed by Intune or Group Policy.
    If not managed, runs Windows Updates using PSWindowsUpdate.
    Generic script - company name passed as parameter.
    Works for multiple companies using the same script file.

    Log saved to company InstallLogs folder - uploaded to R2 by scheduled task.
    Pure ASCII only - safe for 32-bit PowerShell and code page 437.

.PARAMETER CompanyName
    Company name used for log folder path and scheduled task name.
    Must match the value used in Initialize-Device.ps1 for that company.
    Default: DefaultCompany

.EXAMPLE
    .\Check-WindowsUpdate.ps1 -CompanyName "MingLLC"
    .\Check-WindowsUpdate.ps1 -CompanyName "TJBC"

.NOTES
    Designed to run as SYSTEM from a scheduled task.
    Called by the wrapper script Run-WindowsUpdateCheck.ps1.

.VERSION HISTORY
    1.0 - Original release
#>

# ==============================================================================
# PARAMETER - accept company name from caller
# Default set so script can also run standalone without a parameter
# ==============================================================================
param(
    [string]$CompanyName = "DefaultCompany"
)

# ==============================================================================
# CONFIGURATION - do not change
# CompanyName comes from parameter above
# ==============================================================================
$AppName        = "WindowsUpdate"
$AppVersion     = "1.0"
$TaskName       = "$CompanyName-LogUploader"
# ==============================================================================

# ==============================================================================
# DERIVED PATHS - do not change
# ==============================================================================
$LogFolder      = "C:\ProgramData\$CompanyName\InstallLogs"
$LogFileName    = "$($AppName)_$($env:COMPUTERNAME)_$($env:USERNAME)_$($AppVersion)_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$LogPath        = "$LogFolder\$LogFileName"
# ==============================================================================

$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ==============================================================================
# BOOTSTRAP - create log folder before first Write-Log call
# ==============================================================================
if (-not (Test-Path $LogFolder)) {
    New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
}

# ==============================================================================
# LOGGING
# ==============================================================================
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $Entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    Write-Host $Entry
    Add-Content -Path $LogPath -Value $Entry -ErrorAction SilentlyContinue
}

function Exit-AndUpload {
    param([int]$Code)
    Write-Log "Triggering log upload - this is the last log entry."
    Write-Log "--------------------------------------------------------"
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Start-ScheduledTask -TaskName $TaskName
    } else {
        Write-Log "WARNING: Task $TaskName not found - log uploads on next scheduled run." "WARNING"
    }
    exit $Code
}

# ==============================================================================
# CHECK FUNCTIONS
# ==============================================================================
function Test-PSWindowsUpdateModule {
    Write-Log "Checking for PSWindowsUpdate module..."
    $module = Get-Module -ListAvailable -Name PSWindowsUpdate
    if (-not $module) {
        Write-Log "  PSWindowsUpdate not found - installing..." "WARNING"
        try {
            Install-Module -Name PSWindowsUpdate -Force -AcceptLicense -Scope CurrentUser
            Import-Module PSWindowsUpdate
            Write-Log "  PSWindowsUpdate installed successfully."
            return $true
        } catch {
            Write-Log "  Failed to install PSWindowsUpdate: $_" "ERROR"
            return $false
        }
    } else {
        Import-Module PSWindowsUpdate
        Write-Log "  PSWindowsUpdate v$($module.Version) found."
        return $true
    }
}

function Test-IntuneEnrollment {
    Write-Log "Checking Intune enrollment..."
    $intuneEnrolled = $false

    $mdmProviderIDs = @(
        "MS DM Server",
        "Microsoft Device Management",
        "MS MDM Server",
        "ConfigMgr",
        "AutoPilot"
    )

    try {
        $allEnrollments = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Enrollments\*" -ErrorAction SilentlyContinue
        foreach ($enrollment in $allEnrollments) {
            if ($enrollment.ProviderId -and $enrollment.ProviderId -in $mdmProviderIDs) {
                $intuneEnrolled = $true
                Write-Log "  MDM enrollment found: $($enrollment.ProviderId)"
                if ($enrollment.UPN) { Write-Log "  Enrolled user: $($enrollment.UPN)" }
            }
        }
    } catch {
        Write-Log "  Error checking MDM enrollment: $_" "ERROR"
    }

    $intuneService = Get-Service -Name "IntuneManagementExtension" -ErrorAction SilentlyContinue
    if ($intuneService) {
        $intuneEnrolled = $true
        Write-Log "  Intune Management Extension found - Status: $($intuneService.Status)"
    }

    $autopilotInfo = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Provisioning\Diagnostics\Autopilot" -ErrorAction SilentlyContinue
    if ($autopilotInfo -and $autopilotInfo.CloudAssignedOobeConfig) {
        Write-Log "  Autopilot enrollment detected."
    }

    if ($intuneEnrolled) {
        Write-Log "  Result: Device IS enrolled in Intune/MDM."
    } else {
        Write-Log "  Result: Device is NOT enrolled in Intune/MDM." "WARNING"
    }

    return $intuneEnrolled
}

function Test-GroupPolicyManagement {
    Write-Log "Checking Group Policy management..."
    $gpManaged = $false

    $gpPaths = @(
        @{Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate";    Name = "WUServer";    Desc = "WSUS Server configured"},
        @{Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate";    Name = "WUStatusServer"; Desc = "WSUS Status Server configured"},
        @{Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"; Name = "UseWUServer"; Desc = "Use WSUS Server enabled"},
        @{Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"; Name = "NoAutoUpdate"; Desc = "Automatic Updates disabled by policy"},
        @{Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"; Name = "AUOptions";   Desc = "Automatic Update options configured"}
    )

    foreach ($gp in $gpPaths) {
        try {
            $value = Get-ItemProperty -Path $gp.Path -Name $gp.Name -ErrorAction SilentlyContinue
            if ($value) {
                $gpManaged = $true
                Write-Log "  GP setting found - $($gp.Desc): $($value.($gp.Name))"
            }
        } catch { }
    }

    try {
        $cs = Get-CimInstance -Class Win32_ComputerSystem -ErrorAction SilentlyContinue
        if (-not $cs) { $cs = Get-WmiObject -Class Win32_ComputerSystem -ErrorAction SilentlyContinue }
        if ($cs.PartOfDomain) {
            Write-Log "  Device is domain joined to: $($cs.Domain)"
        } else {
            Write-Log "  Device is not domain joined."
        }
    } catch {
        Write-Log "  Error checking domain membership: $_" "ERROR"
    }

    if ($gpManaged) {
        Write-Log "  Result: Windows Update IS managed by Group Policy."
    } else {
        Write-Log "  Result: Windows Update is NOT managed by Group Policy." "WARNING"
    }

    return $gpManaged
}

function Test-WindowsUpdateForBusiness {
    Write-Log "Checking Windows Update for Business..."
    $wufbConfigured = $false

    $wufbPaths = @(
        @{Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"; Name = "DeferFeatureUpdatesPeriodInDays"; Desc = "Feature updates deferred"},
        @{Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"; Name = "DeferQualityUpdatesPeriodInDays"; Desc = "Quality updates deferred"},
        @{Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"; Name = "BranchReadinessLevel"; Desc = "Branch readiness level set"}
    )

    foreach ($wufb in $wufbPaths) {
        try {
            $value = Get-ItemProperty -Path $wufb.Path -Name $wufb.Name -ErrorAction SilentlyContinue
            if ($value) {
                $wufbConfigured = $true
                Write-Log "  WUfB setting found - $($wufb.Desc): $($value.($wufb.Name))"
            }
        } catch { }
    }

    if ($wufbConfigured) {
        Write-Log "  Result: Windows Update for Business IS configured."
    } else {
        Write-Log "  Result: Windows Update for Business is NOT configured." "WARNING"
    }

    return $wufbConfigured
}

function Start-WindowsUpdateProcess {
    Write-Log "--------------------------------------------------------"
    Write-Log "Starting Windows Update process via PSWindowsUpdate..."
    Write-Log "--------------------------------------------------------"

    try {
        Write-Log "Checking for available updates..."
        $updates = Get-WindowsUpdate -MicrosoftUpdate -Verbose

        if ($updates) {
            Write-Log "Found $($updates.Count) update(s):"
            $updates | ForEach-Object {
                Write-Log "  $($_.Title) - $([math]::Round($_.Size/1MB,2)) MB - KB: $($_.KB)"
            }
            Write-Log "Installing updates (reboot suppressed)..."
            Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot -Verbose
            Write-Log "Update installation complete."
        } else {
            Write-Log "No updates available - system is up to date."
        }
    } catch {
        Write-Log "Error during update process: $_" "ERROR"
    }
}

# ==============================================================================
# MAIN
# ==============================================================================
Write-Log "--------------------------------------------------------"
Write-Log "Windows Update Management Script v$AppVersion"
Write-Log "--------------------------------------------------------"
Write-Log "Company   : $CompanyName"
Write-Log "Computer  : $env:COMPUTERNAME"
Write-Log "User      : $env:USERNAME"
Write-Log "Time      : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Log "Log file  : $LogPath"
Write-Log "--------------------------------------------------------"
Write-Log ""

# Check admin rights
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Log "ERROR: Script must run as Administrator." "ERROR"
    Exit-AndUpload -Code 1
}
Write-Log "Administrator privileges confirmed."
Write-Log ""

# Run checks
$intuneManaged = Test-IntuneEnrollment
Write-Log ""
$gpManaged     = Test-GroupPolicyManagement
Write-Log ""
$wufbManaged   = Test-WindowsUpdateForBusiness
Write-Log ""

# Summary and action
Write-Log "--------------------------------------------------------"
Write-Log "SUMMARY"
Write-Log "--------------------------------------------------------"

if ($intuneManaged -or $gpManaged -or $wufbManaged) {
    Write-Log "Windows Update IS MANAGED by:"
    if ($intuneManaged) { Write-Log "  - Intune / MDM" }
    if ($gpManaged)     { Write-Log "  - Group Policy" }
    if ($wufbManaged)   { Write-Log "  - Windows Update for Business" }
    Write-Log "Skipping manual update - organisation management is handling updates."
} else {
    Write-Log "Windows Update is NOT managed - running PSWindowsUpdate..."
    if (Test-PSWindowsUpdateModule) {
        Start-WindowsUpdateProcess
    } else {
        Write-Log "Cannot proceed - PSWindowsUpdate module unavailable." "ERROR"
    }
}

Write-Log ""
Write-Log "End Time  : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Log "--------------------------------------------------------"

Exit-AndUpload -Code 0