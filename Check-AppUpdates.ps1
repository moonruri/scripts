<#
.SYNOPSIS
    Check-AppUpdates.ps1
    Updates all installed applications using winget.
    Generic script - company name passed as parameter.
    Works for multiple companies using the same script file.

    Runs as SYSTEM - includes winget bootstrap for headless/no-profile context.
    Log saved to company InstallLogs folder - uploaded to R2 by scheduled task.
    Pure ASCII only - safe for 32-bit PowerShell and code page 437.

.PARAMETER CompanyName
    Company name used for log folder path and scheduled task name.
    Must match the value used in Initialize-Device.ps1 for that company.
    Default: DefaultCompany

.EXAMPLE
    .\Check-AppUpdates.ps1 -CompanyName "MingLLC"
    .\Check-AppUpdates.ps1 -CompanyName "TJBC"

.NOTES
    Designed to run as SYSTEM where no user profile may be loaded.
    Requires winget (App Installer) to be installed on the device.
    Deploy via Intune or call from scheduled task wrapper.

.VERSION HISTORY
    1.0 - Initial release
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
$AppName        = "AppUpdate"
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
# WINGET BOOTSTRAP
#
# Problem: winget is a user-context app installed per-user via the Microsoft
# Store. When running as SYSTEM, the user profile may not be loaded, so winget
# may not be in PATH.
#
# Solution: Search for winget.exe directly in WindowsApps folder which is the
# system-wide installation location. Sort by LastWriteTime to get newest version.
#
# Note: Even in SYSTEM context, winget can run if called with full path.
# Using --scope machine ensures packages install machine-wide for all users.
# ==============================================================================
function Get-WingetPath {
    Write-Log "Locating winget executable..."

    # Check 1 - winget in PATH (works if user profile loaded or system-wide install)
    $InPath = Get-Command winget -ErrorAction SilentlyContinue
    if ($InPath) {
        Write-Log "  Found in PATH: $($InPath.Source)"
        return $InPath.Source
    }

    # Check 2 - WindowsApps folder (most common location for SYSTEM context)
    $WingetSearch = Get-ChildItem `
        -Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller*_x64__8wekyb3d8bbwe\winget.exe" `
        -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if ($WingetSearch) {
        Write-Log "  Found in WindowsApps: $($WingetSearch.FullName)"
        return $WingetSearch.FullName
    }

    # Check 3 - alternate bundled path
    $AltPath = "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe\winget.exe"
    if (Test-Path $AltPath) {
        Write-Log "  Found at alternate path: $AltPath"
        return $AltPath
    }

    Write-Log "  winget not found in any known location." "WARNING"
    return $null
}

# ==============================================================================
# RUN WINGET WITH OUTPUT CAPTURE
# Captures stdout and stderr line by line into the log
# Returns the process exit code
# ==============================================================================
function Invoke-Winget {
    param(
        [string]$WingetPath,
        [string[]]$Arguments
    )

    Write-Log "  Running: winget $($Arguments -join ' ')"

    try {
        $TempOut = "$env:TEMP\winget_out_$(Get-Random).txt"
        $TempErr = "$env:TEMP\winget_err_$(Get-Random).txt"

        $Process = Start-Process `
            -FilePath             $WingetPath `
            -ArgumentList         $Arguments `
            -Wait `
            -NoNewWindow `
            -PassThru `
            -RedirectStandardOutput $TempOut `
            -RedirectStandardError  $TempErr

        # Log stdout line by line - filter empty lines and progress bar artifacts
        if (Test-Path $TempOut) {
            $OutLines = Get-Content $TempOut -ErrorAction SilentlyContinue
            if ($OutLines) {
                foreach ($Line in $OutLines) {
                    $Clean = $Line.Trim()
                    if ($Clean -and $Clean -notmatch "^[-\\|/]+$" -and $Clean.Length -gt 2) {
                        Write-Log "    $Clean"
                    }
                }
            }
            Remove-Item $TempOut -Force -ErrorAction SilentlyContinue
        }

        # Log stderr if present
        if (Test-Path $TempErr) {
            $ErrLines = Get-Content $TempErr -ErrorAction SilentlyContinue
            if ($ErrLines) {
                foreach ($Line in $ErrLines) {
                    $Clean = $Line.Trim()
                    if ($Clean -and $Clean.Length -gt 2) {
                        Write-Log "    STDERR: $Clean" "WARNING"
                    }
                }
            }
            Remove-Item $TempErr -Force -ErrorAction SilentlyContinue
        }

        Write-Log "  Exit code: $($Process.ExitCode)"
        return $Process.ExitCode

    } catch {
        Write-Log "  ERROR running winget: $($_.Exception.Message)" "ERROR"
        return 1
    }
}

# ==============================================================================
# START
# ==============================================================================
Write-Log "--------------------------------------------------------"
Write-Log "Application Update via Winget v$AppVersion"
Write-Log "--------------------------------------------------------"
Write-Log "Company   : $CompanyName"
Write-Log "Computer  : $env:COMPUTERNAME"
Write-Log "User      : $env:USERNAME"
Write-Log "Time      : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Log "Log file  : $LogPath"
Write-Log "--------------------------------------------------------"
Write-Log ""

# ==============================================================================
# STEP 1 - Check admin rights
# ==============================================================================
Write-Log "STEP 1: Checking privileges..."

$Principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $Principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Log "  ERROR: Script must run as Administrator or SYSTEM." "ERROR"
    Exit-AndUpload -Code 1
}
Write-Log "  Administrator privileges confirmed."
Write-Log "STEP 1: Complete."
Write-Log ""

# ==============================================================================
# STEP 2 - Locate winget
# ==============================================================================
Write-Log "STEP 2: Locating winget..."

$Winget = Get-WingetPath

if (-not $Winget) {
    Write-Log "  ERROR: winget not found." "ERROR"
    Write-Log "  Ensure App Installer (Microsoft.DesktopAppInstaller) is installed." "ERROR"
    Write-Log "  In Intune: deploy Microsoft Store app App Installer to all devices." "ERROR"
    Exit-AndUpload -Code 1
}

if (-not (Test-Path $Winget)) {
    Write-Log "  ERROR: winget path found but file not accessible: $Winget" "ERROR"
    Exit-AndUpload -Code 1
}

Write-Log "  winget path: $Winget"
Write-Log "STEP 2: Complete."
Write-Log ""

# ==============================================================================
# STEP 3 - Refresh winget sources
#
# Using source update not source reset --force
# source reset destroys custom sources - source update just refreshes
# ==============================================================================
Write-Log "STEP 3: Refreshing winget sources..."

$SourceExit = Invoke-Winget -WingetPath $Winget -Arguments @(
    "source", "update",
    "--accept-source-agreements"
)

if ($SourceExit -ne 0) {
    Write-Log "  WARNING: Source update returned exit code $SourceExit" "WARNING"
    Write-Log "  Continuing anyway - source may still be usable." "WARNING"
} else {
    Write-Log "  Sources refreshed successfully."
}

Write-Log "STEP 3: Complete."
Write-Log ""

# ==============================================================================
# STEP 4 - List available upgrades before running
# Logged so you can see what was available even if upgrade had issues
# ==============================================================================
Write-Log "STEP 4: Checking available upgrades..."

Invoke-Winget -WingetPath $Winget -Arguments @(
    "upgrade",
    "--accept-source-agreements",
    "--include-unknown"
) | Out-Null

Write-Log "STEP 4: Complete."
Write-Log ""

# ==============================================================================
# STEP 5 - Run upgrades
#
# Flags:
#   --all                        = upgrade all installed packages
#   --silent                     = no UI prompts
#   --force                      = override installer checks
#   --scope machine              = install machine-wide (critical for SYSTEM)
#   --accept-package-agreements  = auto-accept licenses
#   --accept-source-agreements   = auto-accept source agreements
#   --include-unknown            = include packages with unknown versions
#
# Exit codes:
#   0            = success
#   -1978335189  = no applicable update found
#   -1978335212  = no packages found
#   other        = partial success or error - check log output above
# ==============================================================================
Write-Log "STEP 5: Running upgrades..."

$UpgradeExit = Invoke-Winget -WingetPath $Winget -Arguments @(
    "upgrade",
    "--all",
    "--silent",
    "--force",
    "--scope", "machine",
    "--accept-package-agreements",
    "--accept-source-agreements",
    "--include-unknown"
)

switch ($UpgradeExit) {
    0             { Write-Log "  Upgrades completed successfully." }
    -1978335189   { Write-Log "  No applicable upgrades found - all apps are up to date." }
    -1978335212   { Write-Log "  No packages found matching upgrade criteria." }
    default       {
        Write-Log "  Upgrade completed with exit code: $UpgradeExit" "WARNING"
        Write-Log "  Some packages may not have upgraded - check log output above." "WARNING"
    }
}

Write-Log "STEP 5: Complete."
Write-Log ""

# ==============================================================================
# DONE
# ==============================================================================
Write-Log "--------------------------------------------------------"
Write-Log "Application update COMPLETE."
Write-Log "Company   : $CompanyName"
Write-Log "End Time  : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Log "Check Wasabi/R2 intune-logs/AppUpdate/ for this device log."
Write-Log "--------------------------------------------------------"

Exit-AndUpload -Code 0