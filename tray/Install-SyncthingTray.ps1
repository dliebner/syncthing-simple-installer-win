#Requires -Version 5.1
<#
.SYNOPSIS
    Installs (or removes) the "Syncthing Monitor" tray icon for the current user.

.DESCRIPTION
    - Copies SyncthingTray.ps1, launch-tray.vbs and this script into the Syncthing
      install folder (next to syncthing.exe), so the shortcuts point somewhere stable
    - Adds "Syncthing Monitor" shortcuts to the Startup folder (auto-start at logon),
      the Start menu and the desktop (so it can be relaunched if the icon goes missing)
    - Starts the monitor now, restarting it if it is already running (upgrade case)
    - Does NOT touch Syncthing itself

    Install-Syncthing.ps1 runs this automatically. Run it by hand to add the tray
    icon to an existing install, or with -Uninstall to remove it again.

.PARAMETER InstallDir
    Syncthing install folder the tray files are copied to.
    Defaults to %LOCALAPPDATA%\Programs\Syncthing (same as Install-Syncthing.ps1).

.PARAMETER Uninstall
    Stop the monitor and remove its shortcuts instead of installing.
    The tray files in InstallDir are left in place so it can be reinstalled later.

.PARAMETER NoPause
    Don't wait for Enter at the end, and rethrow errors to the caller.
    Used when this script is run from Install-Syncthing.ps1.
#>

param(
    [string]$InstallDir = "$env:LOCALAPPDATA\Programs\Syncthing",
    [switch]$Uninstall,
    [switch]$NoPause
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ─────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────

$ShortcutName  = "Syncthing Monitor.lnk"
$IconFileName  = "SyncthingMonitor.ico"   # generated at install time by SyncthingTray.ps1 -ExportIcon
$TrayFiles     = @("SyncthingTray.ps1", "launch-tray.vbs", "Install-SyncthingTray.ps1")
$ShortcutDirs  = @(
    [Environment]::GetFolderPath('Startup'),    # auto-start at logon
    [Environment]::GetFolderPath('Programs'),   # Start menu (searchable)
    [Environment]::GetFolderPath('Desktop')
)
# Must match $ExitEventName in SyncthingTray.ps1.
$ExitEventName = "SyncthingTrayMonitor_Exit_$env:USERNAME"

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

function Get-TrayProcess {
    # The launcher runs: powershell.exe ... -File "<dir>\SyncthingTray.ps1". The leading
    # backslash keeps this from matching Install-SyncthingTray.ps1 (i.e. ourselves).
    Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" |
        Where-Object { $_.ProcessId -ne $PID -and $_.CommandLine -and $_.CommandLine -like '*\SyncthingTray.ps1*' }
}

function Stop-TrayMonitor {
    # Ask a running instance to exit cleanly (so its icon is removed), then make
    # sure it is actually gone before we carry on.
    try {
        $evt = [System.Threading.EventWaitHandle]::OpenExisting($ExitEventName)
        [void]$evt.Set()
        $evt.Dispose()
    } catch { }   # not running (or an older build without the exit event)

    $deadline = (Get-Date).AddSeconds(5)
    while ((Get-Date) -lt $deadline) {
        if (-not (Get-TrayProcess)) { return }
        Start-Sleep -Milliseconds 250
    }
    Get-TrayProcess | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
}

try {

# ─────────────────────────────────────────────
# UNINSTALL
# ─────────────────────────────────────────────

if ($Uninstall) {
    Write-Step "Stopping Syncthing Monitor..."
    Stop-TrayMonitor
    Write-Success "Stopped."

    Write-Step "Removing shortcuts..."
    foreach ($dir in $ShortcutDirs) {
        $lnk = Join-Path $dir $ShortcutName
        if (Test-Path $lnk) {
            Remove-Item $lnk -Force
            Write-Success "Removed $lnk"
        }
    }

    Write-Host ""
    Write-Host "Syncthing Monitor removed. It will no longer start at logon." -ForegroundColor Green
    Write-Host "The tray files are still in $InstallDir (re-run this script to reinstall,"
    Write-Host "or delete them if you no longer want them)."
    return
}

# ─────────────────────────────────────────────
# STEP 1: COPY FILES INTO THE INSTALL FOLDER
# ─────────────────────────────────────────────

$source = $PSScriptRoot
foreach ($f in $TrayFiles) {
    if (-not (Test-Path (Join-Path $source $f))) {
        throw "Can't find $f next to this script. Keep the tray files together."
    }
}

Write-Step "Copying tray files to $InstallDir..."

if (-not (Test-Path $InstallDir)) {
    New-Item -ItemType Directory -Path $InstallDir | Out-Null
}

$sameDir = (Resolve-Path $source).Path.TrimEnd('\') -eq (Resolve-Path $InstallDir).Path.TrimEnd('\')
if ($sameDir) {
    Write-Success "Already running from $InstallDir, nothing to copy."
} else {
    foreach ($f in $TrayFiles) {
        Copy-Item -Path (Join-Path $source $f) -Destination (Join-Path $InstallDir $f) -Force
    }
    Write-Success "Copied $($TrayFiles -join ', ')"
}

# Clear the mark-of-the-web from files that came out of a downloaded ZIP, so
# Windows doesn't show a security prompt when the shortcut runs at logon.
Unblock-File -Path ($TrayFiles | ForEach-Object { Join-Path $InstallDir $_ }) -ErrorAction SilentlyContinue

# ─────────────────────────────────────────────
# STEP 2: SHORTCUT ICON
# ─────────────────────────────────────────────

# Ask the tray script to draw its own "running" icon (folder + green badge) as a
# multi-size .ico, so the shortcuts look exactly like the tray. Runs in a child
# process so its own settings/strict-mode don't leak into this one.
Write-Step "Generating shortcut icon..."

$iconPath     = Join-Path $InstallDir $IconFileName
$iconLocation = "shell32.dll,3"    # plain folder, used only if generation fails
& powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass `
    -File (Join-Path $InstallDir "SyncthingTray.ps1") -ExportIcon $iconPath
if ($LASTEXITCODE -eq 0 -and (Test-Path $iconPath)) {
    $iconLocation = "$iconPath,0"
    Write-Success "Created $iconPath"
} else {
    Write-Host "    WARN: Couldn't generate the icon; shortcuts will use the plain folder icon." -ForegroundColor Yellow
}

# ─────────────────────────────────────────────
# STEP 3: SHORTCUTS
# ─────────────────────────────────────────────

Write-Step "Creating 'Syncthing Monitor' shortcuts..."

$launcher = Join-Path $InstallDir "launch-tray.vbs"
$ws = New-Object -ComObject WScript.Shell
foreach ($dir in $ShortcutDirs) {
    $lnk = Join-Path $dir $ShortcutName
    $sc  = $ws.CreateShortcut($lnk)
    $sc.TargetPath       = "wscript.exe"          # runs the .vbs, which launches the tray hidden
    $sc.Arguments        = """$launcher"""
    $sc.WorkingDirectory = $InstallDir
    $sc.IconLocation     = $iconLocation          # same folder-with-badge icon as the tray
    $sc.Description      = "Shows whether Syncthing is running, in the system tray."
    $sc.Save()
    Write-Success $lnk
}

# ─────────────────────────────────────────────
# STEP 4: START (OR RESTART) IT NOW
# ─────────────────────────────────────────────

Write-Step "Starting Syncthing Monitor..."

# Stop any running instance first: its single-instance guard would otherwise
# keep the old copy running (matters when re-running this to upgrade).
if (Get-TrayProcess) { Stop-TrayMonitor }
Start-Process "wscript.exe" -ArgumentList """$launcher"""
Write-Success "Running now, and will start automatically at logon."

Write-Host ""
Write-Host "If the tray icon ever disappears, reopen 'Syncthing Monitor' from the Start menu or desktop."
Write-Host "To remove it later: run this script with -Uninstall (or delete the three shortcuts above)."

} catch {

    if ($NoPause) { throw }   # let the calling installer report it
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""

} finally {

    if (-not $NoPause) { Read-Host "Press Enter to exit..." }

}
