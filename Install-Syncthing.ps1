#Requires -Version 5.1
<#
.SYNOPSIS
    Installs and configures Syncthing as a scheduled task on Windows.

.DESCRIPTION
    - Downloads the latest Syncthing release directly from GitHub
    - Installs to %LOCALAPPDATA%\Programs\Syncthing (allows Syncthing to self-update)
    - Creates a Task Scheduler task to start Syncthing at user logon
    - Adds a Windows Firewall rule for Syncthing
    - Starts Syncthing immediately after installation

.NOTES
    Run this script as the user who will be running Syncthing.
    Firewall rule creation requires admin privileges (script will prompt via UAC).

    Run Command: Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass; .\Install-Syncthing.ps1

.PARAMETER InstallDir
    Directory to install Syncthing. Defaults to %LOCALAPPDATA%\Syncthing.
    Must be writable by the current user to allow Syncthing self-updates.

.PARAMETER GuiPort
    Port for Syncthing's web UI. Defaults to 8384.

.PARAMETER StartupDelay
    Seconds to delay Syncthing startup after logon, to avoid startup congestion.
    Defaults to 30 seconds.
#>

param(
    [string]$InstallDir = "$env:LOCALAPPDATA\Programs\Syncthing",
    [int]$GuiPort = 8384,
    [int]$StartupDelay = 30
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ─────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────

$TaskNameStart    = "Syncthing - Start at Logon ($env:USERNAME)"
$FirewallRuleName = "Syncthing ($env:USERNAME)"
$SyncthingExe     = Join-Path $InstallDir "syncthing.exe"
$ConfigPath       = Join-Path $env:LOCALAPPDATA "Syncthing\config.xml"

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

function Write-Step {
    param([string]$Message)
    Write-Host "`n==> $Message" -ForegroundColor Cyan
}

function Write-Success {
    param([string]$Message)
    Write-Host "    OK: $Message" -ForegroundColor Green
}

function Write-Warn {
    param([string]$Message)
    Write-Host "    WARN: $Message" -ForegroundColor Yellow
}


# ─────────────────────────────────────────────
# STEP 1: DOWNLOAD SYNCTHING
# ─────────────────────────────────────────────

Write-Step "Fetching latest Syncthing release from GitHub..."

$releaseApi = "https://api.github.com/repos/syncthing/syncthing/releases/latest"
$release    = Invoke-RestMethod -Uri $releaseApi -UseBasicParsing
$version    = $release.tag_name

# Determine architecture
$isArm = ($env:PROCESSOR_ARCHITECTURE -match 'ARM64') -or ($env:PROCESSOR_ARCHITEW6432 -match 'ARM64')
$arch = if ($isArm) { "arm64" } elseif ([Environment]::Is64BitOperatingSystem) { "amd64" } else { "386" }
$assetName = "syncthing-windows-$arch-$version.zip"
$asset     = $release.assets | Where-Object { $_.name -eq $assetName }

if (-not $asset) {
    throw "Could not find release asset '$assetName' in GitHub release $version."
}

Write-Success "Found Syncthing $version ($assetName)"

# Download to temp
$zipPath = Join-Path $env:TEMP $assetName
Write-Host "    Downloading to $zipPath..."
Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $zipPath -UseBasicParsing
Write-Success "Download complete."

# Stop any existing syncthing process
if (Get-Process "syncthing" -ErrorAction SilentlyContinue) {
    Stop-Process -Name "syncthing" -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
}

# ─────────────────────────────────────────────
# STEP 2: INSTALL
# ─────────────────────────────────────────────

Write-Step "Installing Syncthing to $InstallDir..."

if (-not (Test-Path $InstallDir)) {
    New-Item -ItemType Directory -Path $InstallDir | Out-Null
}

# Extract zip
$extractTemp = Join-Path $env:TEMP "syncthing-extract"
if (Test-Path $extractTemp) { Remove-Item $extractTemp -Recurse -Force }
Expand-Archive -Path $zipPath -DestinationPath $extractTemp -Force

# The zip contains a subdirectory e.g. syncthing-windows-amd64-v1.x.x\
$extractedDir = Get-ChildItem $extractTemp -Directory | Select-Object -First 1
$sourceExe     = Join-Path $extractedDir.FullName "syncthing.exe"

Copy-Item -Path $sourceExe -Destination $SyncthingExe -Force
Write-Success "syncthing.exe installed to $SyncthingExe"

# Cleanup
Remove-Item $zipPath -Force
Remove-Item $extractTemp -Recurse -Force

# ─────────────────────────────────────────────
# STEP 3: FIRST RUN (generate config + API key)
# ─────────────────────────────────────────────

Write-Step "Generating Syncthing configuration..."

# Only do this if config doesn't already exist
if (-not (Test-Path $ConfigPath)) {
    & $SyncthingExe generate | Out-Null
    Write-Success "Configuration generated."
} else {
    Write-Success "Configuration already exists, skipping generation."
}

# ─────────────────────────────────────────────
# STEP 4: UPDATE CONFIG.XML
# ─────────────────────────────────────────────

Write-Step "Setting GUI port in config.xml..."

[xml]$configXml = Get-Content $ConfigPath

# Update the GUI port permanently in the XML
if ($configXml.configuration.gui.address -ne "127.0.0.1:$GuiPort") {
    $configXml.configuration.gui.address = "127.0.0.1:$GuiPort"
    $configXml.Save($ConfigPath)
    Write-Success "Updated config.xml GUI address to 127.0.0.1:$GuiPort"
} else {
    Write-Success "GUI address is already set correctly in config.xml."
}

# ─────────────────────────────────────────────
# STEP 5: TASK SCHEDULER - START AT LOGON
# ─────────────────────────────────────────────

Write-Step "Creating scheduled task: '$TaskNameStart'..."

# Remove existing task if present
if (Get-ScheduledTask -TaskName $TaskNameStart -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $TaskNameStart -Confirm:$false
}

$startAction = New-ScheduledTaskAction `
    -Execute "`"$SyncthingExe`"" `
    -Argument "--no-browser --no-console"

# Trigger: at logon of current user, with startup delay
$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$startTrigger = New-ScheduledTaskTrigger -AtLogOn -User $currentUser
$startTrigger.Delay = "PT${StartupDelay}S"   # ISO 8601 duration

$settings = New-ScheduledTaskSettingsSet `
    -Hidden `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1) `
    -DontStopIfGoingOnBatteries `
    -AllowStartIfOnBatteries

$settings.ExecutionTimeLimit = "PT0S" # No time limit

$principal = New-ScheduledTaskPrincipal `
    -UserId $currentUser `
    -LogonType Interactive `
    -RunLevel Limited

Register-ScheduledTask `
    -TaskName $TaskNameStart `
    -Action $startAction `
    -Trigger $startTrigger `
    -Settings $settings `
    -Principal $principal `
    -Description "Starts Syncthing file synchronization at user logon." | Out-Null

Write-Success "Start task created."

# ─────────────────────────────────────────────
# STEP 6: FIREWALL RULE
# ─────────────────────────────────────────────

Write-Step "Adding Windows Firewall rule for Syncthing..."

# Firewall changes require admin — attempt elevation if needed
$isAdmin = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent() `
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if ($isAdmin) {
    # Remove existing rule if present
    Remove-NetFirewallRule -DisplayName $FirewallRuleName -ErrorAction SilentlyContinue

    New-NetFirewallRule `
        -DisplayName $FirewallRuleName `
        -Direction Inbound `
        -Program $SyncthingExe `
        -Action Allow `
        -Profile @("Private","Domain") | Out-Null

    Write-Success "Firewall rule created."
} else {
    Write-Warn "Prompting for Admin privileges to add the Firewall Rule..."
    $safeExe = $SyncthingExe -replace "'", "''"
    $safeRuleName = $FirewallRuleName -replace "'", "''"
    $fwScript = "Remove-NetFirewallRule -DisplayName '$safeRuleName' -ErrorAction SilentlyContinue; New-NetFirewallRule -DisplayName '$safeRuleName' -Direction Inbound -Program '$safeExe' -Action Allow -Profile @('Private','Domain')"
    $encoded  = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($fwScript))
    try {
        Start-Process powershell -Verb RunAs -ArgumentList "-NoProfile -WindowStyle Hidden -EncodedCommand $encoded" -Wait
        Write-Success "Firewall rule creation finished."
    } catch {
        Write-Warn "UAC prompt was canceled. Firewall rule was not created."
    }
}

# ─────────────────────────────────────────────
# DONE
# ─────────────────────────────────────────────

Write-Step "Starting Syncthing..."
Start-ScheduledTask -TaskName $TaskNameStart
Write-Success "Syncthing is now running in the background."

Write-Host ""
Write-Host "─────────────────────────────────────────────" -ForegroundColor Cyan
Write-Host "  Syncthing installation complete!" -ForegroundColor Green
Write-Host "─────────────────────────────────────────────" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Installed to : $SyncthingExe"
Write-Host "  Web UI       : http://localhost:$GuiPort"
Write-Host "  Start task   : $TaskNameStart (delay: ${StartupDelay}s after logon)"
Write-Host ""
Write-Host "  Next steps:"
Write-Host "  1. Open http://localhost:$GuiPort to configure Syncthing"
Write-Host ""
Write-Host "  To uninstall:"
Write-Host "  - Delete $InstallDir"
Write-Host "  - Remove scheduled task '$TaskNameStart'"
Write-Host "  - Remove firewall rule '$FirewallRuleName'"
Write-Host "  - Optional: Delete config and db at $env:LOCALAPPDATA\Syncthing"
Write-Host ""
