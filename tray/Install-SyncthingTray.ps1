#Requires -Version 5.1
<#
.SYNOPSIS
    Installs (or removes) the "Syncthing Monitor" tray icon for the current user.

.DESCRIPTION
    - Copies SyncthingTray.ps1 and this script into the Syncthing install folder
      (next to syncthing.exe), so nothing points at a Downloads folder
    - Registers a scheduled task that starts the tray icon at logon, the same way
      Install-Syncthing.ps1 starts Syncthing itself. No Startup-folder shortcut and
      no .vbs launcher: allow-list antivirus (PC Matic, for one) blocks shortcuts
      that launch script interpreters, which silently kills that kind of autostart.
    - Adds "Syncthing Monitor" shortcuts to the Start menu and the desktop that
      trigger the same task, so the icon can be relaunched if it ever goes missing
    - Starts the monitor now, restarting it if it is already running (upgrade case)
    - Does NOT touch Syncthing itself

    Install-Syncthing.ps1 runs this automatically. Run it by hand to add the tray
    icon to an existing install, or with -Uninstall to remove it again.

.PARAMETER InstallDir
    Syncthing install folder the tray files are copied to.
    Defaults to %LOCALAPPDATA%\Programs\Syncthing (same as Install-Syncthing.ps1).

.PARAMETER Uninstall
    Stop the monitor, remove its scheduled task and shortcuts instead of installing.
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

$TaskName      = "Syncthing Monitor ($env:USERNAME)"
$ShortcutName  = "Syncthing Monitor.lnk"
$IconFileName  = "SyncthingMonitor.ico"   # generated at install time by SyncthingTray.ps1 -ExportIcon
$TrayFiles     = @("SyncthingTray.ps1", "Install-SyncthingTray.ps1")
$ShortcutDirs  = @(
    [Environment]::GetFolderPath('Programs'),   # Start menu (searchable)
    [Environment]::GetFolderPath('Desktop')
)
# Older versions put an autostart shortcut here; it is removed on install/uninstall.
$LegacyStartupLnk = Join-Path ([Environment]::GetFolderPath('Startup')) $ShortcutName

$PowerShellExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$TrayArguments = "-NoProfile -NonInteractive -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$(Join-Path $InstallDir 'SyncthingTray.ps1')`""

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

function Write-Warn {
    param([string]$Message)
    Write-Host "    WARN: $Message" -ForegroundColor Yellow
}

function Get-TrayProcess {
    # The task runs: powershell.exe ... -File "<dir>\SyncthingTray.ps1". The leading
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

function Remove-Shortcuts {
    foreach ($lnk in (@($ShortcutDirs | ForEach-Object { Join-Path $_ $ShortcutName }) + $LegacyStartupLnk)) {
        if (Test-Path $lnk) {
            Remove-Item $lnk -Force
            Write-Success "Removed $lnk"
        }
    }
}

try {

# ─────────────────────────────────────────────
# UNINSTALL
# ─────────────────────────────────────────────

if ($Uninstall) {
    Write-Step "Stopping Syncthing Monitor..."
    Stop-TrayMonitor
    Write-Success "Stopped."

    Write-Step "Removing scheduled task '$TaskName'..."
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Success "Removed."
    } else {
        Write-Success "Not present."
    }

    Write-Step "Removing shortcuts..."
    Remove-Shortcuts

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

# Older versions used a .vbs launcher; tidy it up if it's still there.
Remove-Item (Join-Path $InstallDir "launch-tray.vbs") -Force -ErrorAction SilentlyContinue

# Clear the mark-of-the-web from files that came out of a downloaded ZIP.
Unblock-File -Path ($TrayFiles | ForEach-Object { Join-Path $InstallDir $_ }) -ErrorAction SilentlyContinue

# ─────────────────────────────────────────────
# STEP 2: SHORTCUT ICON
# ─────────────────────────────────────────────

# Ask the tray script to draw its own "running" icon (folder + green badge) as a
# multi-size .ico, so the shortcuts look exactly like the tray. Runs in a child
# process so its own settings/strict-mode don't leak into this one. Written to a
# temp file first so a failure can't clobber a good icon from a previous run.
Write-Step "Generating shortcut icon..."

$iconPath = Join-Path $InstallDir $IconFileName
$iconTemp = "$iconPath.tmp"
& $PowerShellExe -NoProfile -NonInteractive -ExecutionPolicy Bypass `
    -File (Join-Path $InstallDir "SyncthingTray.ps1") -ExportIcon $iconTemp
if ($LASTEXITCODE -eq 0 -and (Test-Path $iconTemp)) {
    Move-Item -Path $iconTemp -Destination $iconPath -Force
    Write-Success "Created $iconPath"
} else {
    Remove-Item $iconTemp -Force -ErrorAction SilentlyContinue
    if (Test-Path $iconPath) {
        Write-Warn "Couldn't regenerate the icon; keeping the existing one."
    } else {
        Write-Warn "Couldn't generate the icon; shortcuts will use the plain folder icon."
    }
}
$iconLocation = if (Test-Path $iconPath) { "$iconPath,0" } else { "shell32.dll,3" }

# ─────────────────────────────────────────────
# STEP 3: SCHEDULED TASK - START AT LOGON
# ─────────────────────────────────────────────

Write-Step "Creating scheduled task: '$TaskName'..."

# Stop a running instance before replacing the task, so the restart below picks
# up the new files (its single-instance guard would otherwise keep the old copy).
if (Get-TrayProcess) { Stop-TrayMonitor }

if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

$action = New-ScheduledTaskAction `
    -Execute $PowerShellExe `
    -Argument $TrayArguments `
    -WorkingDirectory $InstallDir

$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $currentUser

$settings = New-ScheduledTaskSettingsSet `
    -Hidden `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1) `
    -DontStopIfGoingOnBatteries `
    -AllowStartIfOnBatteries `
    -MultipleInstances IgnoreNew

$settings.ExecutionTimeLimit = "PT0S"           # no time limit: the tray runs for the whole session
$settings.IdleSettings.StopOnIdleEnd = $false

$principal = New-ScheduledTaskPrincipal `
    -UserId $currentUser `
    -LogonType Interactive `
    -RunLevel Limited

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Principal $principal `
    -Description "Shows whether Syncthing is running, in the system tray." | Out-Null

Write-Success "Task created."

# ─────────────────────────────────────────────
# STEP 4: SHORTCUTS (relaunch by hand)
# ─────────────────────────────────────────────

Write-Step "Creating 'Syncthing Monitor' shortcuts..."

# The shortcuts just poke the task, so the tray always starts the same way
# (and Explorer never launches an interpreter directly).
Remove-Item $LegacyStartupLnk -Force -ErrorAction SilentlyContinue
$ws = New-Object -ComObject WScript.Shell
foreach ($dir in $ShortcutDirs) {
    $lnk = Join-Path $dir $ShortcutName
    $sc  = $ws.CreateShortcut($lnk)
    $sc.TargetPath       = "$env:SystemRoot\System32\schtasks.exe"
    $sc.Arguments        = "/Run /TN `"$TaskName`""
    $sc.WorkingDirectory = $InstallDir
    $sc.WindowStyle      = 7                      # minimized: schtasks' console never lands on screen
    $sc.IconLocation     = $iconLocation          # same folder-with-badge icon as the tray
    $sc.Description      = "Shows whether Syncthing is running, in the system tray."
    $sc.Save()
    Write-Success $lnk
}

# ─────────────────────────────────────────────
# STEP 5: START IT NOW
# ─────────────────────────────────────────────

Write-Step "Starting Syncthing Monitor..."

# Through the task, so this exercises exactly the path used at logon. If an
# allow-list antivirus blocks the launch, its prompt (if enabled) appears right
# here; we wait for the user to deal with it and retry, so one run of this
# script is enough.
function Start-TrayAndWait {
    Start-ScheduledTask -TaskName $TaskName
    $deadline = (Get-Date).AddSeconds(10)
    while (-not (Get-TrayProcess) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 250 }
    return [bool](Get-TrayProcess)
}

function Write-TaskDiagnostics {
    # Report what Task Scheduler thinks happened, so a block can be told apart
    # from a misconfigured task or a script that exited on its own.
    $info = Get-ScheduledTaskInfo -TaskName $TaskName -ErrorAction SilentlyContinue
    $task = Get-ScheduledTask     -TaskName $TaskName -ErrorAction SilentlyContinue
    if (-not ($info -and $task)) { return }
    $code = "0x{0:X}" -f $info.LastTaskResult
    $hint = switch ($info.LastTaskResult) {
        0          { "the process started and exited immediately (typical of antivirus blocking it, or the script failing at startup)" }
        0x80070005 { "'Access is denied': Task Scheduler was refused when launching powershell.exe, which is what allow-list antivirus (e.g. PC Matic SuperShield) looks like" }
        0x41301    { "Task Scheduler says it is still running, so the tray process may simply not have been found by name" }
        0x41303    { "the task has never run; Task Scheduler didn't launch it at all" }
        0x800710E0 { "'the operator or administrator has refused the request' (task conditions/policy stopped it)" }
        default    { "see Task Scheduler > Task Scheduler Library > '$TaskName' > History" }
    }
    Write-Warn "Task state: $($task.State); last result: $code, i.e. $hint."
}

$started = Start-TrayAndWait
while (-not $started) {
    Write-Warn "The task was triggered but no tray process appeared within 10s."
    Write-TaskDiagnostics
    Write-Host ""
    Write-Host "    If your antivirus just prompted about powershell.exe, choose its 'always allow' option." -ForegroundColor Yellow
    Write-Host "    If it blocked silently, open it and allow powershell.exe running SyncthingTray.ps1 from" -ForegroundColor Yellow
    Write-Host "    $InstallDir (PC Matic: SuperShield > Blocking Notification Method >" -ForegroundColor Yellow
    Write-Host "    'Prompt for Override', then retry here and click 'Always Allow')." -ForegroundColor Yellow
    Write-Host ""
    $answer = Read-Host "    Press Enter to try again, or type S to skip for now"
    if ($answer -match '^[sS]') {
        Write-Warn "Skipped. The task stays registered; run this script again once the launch is allowed."
        break
    }
    $started = Start-TrayAndWait
}
if ($started) {
    Write-Success "Running now, and will start automatically at logon."
}

Write-Host ""
Write-Host "If the tray icon ever disappears, reopen 'Syncthing Monitor' from the Start menu or desktop."
Write-Host "To remove it later: run this script with -Uninstall."

} catch {

    if ($NoPause) { throw }   # let the calling installer report it
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""

} finally {

    if (-not $NoPause) { Read-Host "Press Enter to exit..." }

}
