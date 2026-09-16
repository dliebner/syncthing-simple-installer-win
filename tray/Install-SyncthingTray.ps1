#Requires -Version 5.1
<#
.SYNOPSIS
    Installs (or removes) the "Syncthing Monitor" tray icon for the current user.

.DESCRIPTION
    - Copies SyncthingMonitor.cs and this script into the Syncthing install folder
      (next to syncthing.exe) and compiles the tray icon into SyncthingMonitor.exe
      there, using the C# compiler that ships with Windows (.NET Framework 4.x).
      No build tools, no binary in the repo, and the result is a real Windows
      program: no interpreter, no execution-policy bypass, no console window.
    - Registers a scheduled task that starts it at logon, the same way
      Install-Syncthing.ps1 starts Syncthing itself
    - Adds "Syncthing Monitor" shortcuts to the Start menu and the desktop, so the
      icon can be relaunched if it ever goes missing
    - Starts the monitor now, restarting it if it is already running (upgrade case)
    - Does NOT touch Syncthing itself

    Install-Syncthing.ps1 runs this automatically. Run it by hand to add the tray
    icon to an existing install, or with -Uninstall to remove it again.

    Allow-list antivirus (PC Matic and similar) will block the freshly compiled
    exe until it is allowed once; this script waits for that and retries.

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
$SourceName    = "SyncthingMonitor.cs"
$ExeName       = "SyncthingMonitor.exe"
$IconFileName  = "SyncthingMonitor.ico"   # generated at install time by SyncthingMonitor.exe --export-icon
$TrayFiles     = @($SourceName, "Install-SyncthingTray.ps1")
$ShortcutDirs  = @(
    [Environment]::GetFolderPath('Programs'),   # Start menu (searchable)
    [Environment]::GetFolderPath('Desktop')
)
# Older versions put an autostart shortcut here; it is removed on install/uninstall.
$LegacyStartupLnk = Join-Path ([Environment]::GetFolderPath('Startup')) $ShortcutName

$ExePath   = Join-Path $InstallDir $ExeName
$StampPath = "$ExePath.source.sha256"   # hash of the source the exe was built from

# Must match Config.ExitEventName in SyncthingMonitor.cs.
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
    # The compiled tray, plus any instance of the old PowerShell-based tray
    # (SyncthingTray.ps1) left over from earlier versions.
    @(Get-Process -Name "SyncthingMonitor" -ErrorAction SilentlyContinue) +
    @(Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" |
        Where-Object { $_.ProcessId -ne $PID -and $_.CommandLine -and $_.CommandLine -like '*\SyncthingTray.ps1*' } |
        ForEach-Object { Get-Process -Id $_.ProcessId -ErrorAction SilentlyContinue })
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
    Get-TrayProcess | ForEach-Object { Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue }
    Start-Sleep -Milliseconds 500   # let the exe's file handle go before we overwrite it
}

function Remove-Shortcuts {
    foreach ($lnk in (@($ShortcutDirs | ForEach-Object { Join-Path $_ $ShortcutName }) + $LegacyStartupLnk)) {
        if (Test-Path $lnk) {
            Remove-Item $lnk -Force
            Write-Success "Removed $lnk"
        }
    }
}

function Test-PCMatic {
    # PC Matic / SuperShield is an allow-list ("default-deny") antivirus: it blocks
    # any program it doesn't recognize, and a freshly compiled exe never is. Worse,
    # it blocks unknown programs that are *launched by a script or a scheduled task*
    # as a living-off-the-land defence, and logs the block against the launcher
    # (powershell.exe / Task Scheduler), not against our exe - so no "allow this
    # app?" prompt appears and there's nothing named after our exe to whitelist.
    # Detected via its running processes/services so we can give exact guidance.
    $pat = 'pcmatic|supershield|pcpitstop'
    if (Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -match $pat }) { return $true }
    if (Get-Service -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match $pat -or $_.DisplayName -match 'PC Matic|SuperShield|PCPitstop' }) { return $true }
    return $false
}

function Invoke-WhitelistAssist {
    # Get an allow-list AV to let the exe through. The trick: a human double-click
    # from Explorer runs the exe with explorer.exe as its parent - a user-initiated
    # launch - which is the one path such products treat as "the user meant to run
    # this" and either allow or offer to allow. It also finally logs the block
    # against the exe itself, so it shows up in the AV's list as something to
    # whitelist. Either way, allowing it once is scoped to this one small exe.
    param([string]$Path)

    Write-Host ""
    if (Test-PCMatic) {
        Write-Host "    PC Matic (SuperShield) is running and is blocking $ExeName." -ForegroundColor Yellow
        Write-Host "    It blocks unknown programs that are started by a script or a scheduled task," -ForegroundColor Yellow
        Write-Host "    which is why only powershell.exe shows in its blocked list, not $ExeName," -ForegroundColor Yellow
        Write-Host "    and why no allow prompt appeared. Starting it once by hand fixes that:" -ForegroundColor Yellow
    } else {
        Write-Host "    An allow-list antivirus is blocking $ExeName (a freshly compiled program is" -ForegroundColor Yellow
        Write-Host "    unknown to these products). Starting it once by hand lets you allow it:" -ForegroundColor Yellow
    }
    Write-Host ""
    Write-Host "      1. An Explorer window will open with $ExeName selected." -ForegroundColor Yellow
    Write-Host "      2. Double-click $ExeName." -ForegroundColor Yellow
    Write-Host "      3. When your antivirus notifies or blocks it, choose Allow / Always Allow." -ForegroundColor Yellow
    Write-Host "         (Or open PC Matic > SuperShield, find $ExeName in the recently blocked" -ForegroundColor Yellow
    Write-Host "          list, and whitelist it there - it will be listed by name now.)" -ForegroundColor Yellow
    Write-Host "      It's fine if nothing visible happens: the tray icon has no window." -ForegroundColor Yellow
    Write-Host ""

    try {
        Start-Process explorer.exe -ArgumentList "/select,`"$Path`""
    } catch {
        Write-Warn "Couldn't open Explorer automatically. Open this folder and double-click ${ExeName}:"
        Write-Warn "  $Path"
    }
    [void](Read-Host "    Once you've allowed $ExeName, press Enter to continue")
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

# Leftovers from earlier versions (PowerShell tray + .vbs launcher).
foreach ($old in @("SyncthingTray.ps1", "launch-tray.vbs")) {
    Remove-Item (Join-Path $InstallDir $old) -Force -ErrorAction SilentlyContinue
}

# Clear the mark-of-the-web from files that came out of a downloaded ZIP.
Unblock-File -Path ($TrayFiles | ForEach-Object { Join-Path $InstallDir $_ }) -ErrorAction SilentlyContinue

# ─────────────────────────────────────────────
# STEP 2: COMPILE THE TRAY ICON
# ─────────────────────────────────────────────

Write-Step "Building $ExeName..."

# Only rebuild when the source changed. Every build produces a different file
# hash, and allow-list antivirus keys its permission to that hash, so a pointless
# rebuild would mean a pointless re-prompt.
$srcPath = Join-Path $InstallDir $SourceName
$srcHash = (Get-FileHash -Path $srcPath -Algorithm SHA256).Hash
$upToDate = (Test-Path $ExePath) -and (Test-Path $StampPath) -and
            ((Get-Content -Path $StampPath -Raw).Trim() -eq $srcHash)

if ($upToDate) {
    Write-Success "$ExeName is up to date with $SourceName, not rebuilding."
} else {
    # The exe can't be overwritten while it runs.
    if (Get-TrayProcess) { Stop-TrayMonitor }
    Remove-Item $StampPath -Force -ErrorAction SilentlyContinue

    # Add-Type drives the .NET Framework C# compiler (csc.exe, present on every
    # Windows 10/11). -OutputType WindowsApplication = a GUI exe with no console.
    Add-Type -Path $srcPath `
        -OutputAssembly $ExePath `
        -OutputType WindowsApplication `
        -ReferencedAssemblies System.Windows.Forms, System.Drawing, System.Xml `
        -IgnoreWarnings

    if (-not (Test-Path $ExePath)) { throw "The compiler produced no $ExeName." }
    Set-Content -Path $StampPath -Value $srcHash -Encoding ASCII
    Write-Success "Built $ExePath"
}

# ─────────────────────────────────────────────
# STEP 3: SCHEDULED TASK - START AT LOGON
# ─────────────────────────────────────────────

Write-Step "Creating scheduled task: '$TaskName'..."

# Stop a running instance before replacing the task, so the restart below picks
# up the new build (its single-instance guard would otherwise keep the old copy).
if (Get-TrayProcess) { Stop-TrayMonitor }

if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

$action = New-ScheduledTaskAction -Execute $ExePath -WorkingDirectory $InstallDir

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
# STEP 4: START IT NOW
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
    # from a misconfigured task or a program that exited on its own.
    $info = Get-ScheduledTaskInfo -TaskName $TaskName -ErrorAction SilentlyContinue
    $task = Get-ScheduledTask     -TaskName $TaskName -ErrorAction SilentlyContinue
    if (-not ($info -and $task)) { return }
    $code = "0x{0:X}" -f $info.LastTaskResult
    $hint = switch ($info.LastTaskResult) {
        0          { "the process started and exited immediately (antivirus blocking it, or it failed at startup)" }
        0x80070005 { "'Access is denied': Task Scheduler was refused when launching $ExeName, which is what allow-list antivirus (e.g. PC Matic SuperShield) looks like" }
        0x41301    { "Task Scheduler says it is still running, so the process may simply not have been found by name" }
        0x41303    { "the task has never run; Task Scheduler didn't launch it at all" }
        0x800710E0 { "'the operator or administrator has refused the request' (task conditions/policy stopped it)" }
        default    { "see Task Scheduler > Task Scheduler Library > '$TaskName' > History" }
    }
    Write-Warn "Task state: $($task.State); last result: $code, i.e. $hint."
}

$started = Start-TrayAndWait
while (-not $started) {
    Write-Warn "The task was triggered but $ExeName did not stay running for 10s."
    Write-TaskDiagnostics

    # A task/script-launched start is exactly what an allow-list AV refuses, so
    # walk the user through starting it once by hand (Explorer double-click) to
    # get it allowed. Once the file is whitelisted, the task launch below works
    # too, because the allow is on the exe, not on whatever launches it.
    Invoke-WhitelistAssist -Path $ExePath

    $started = Start-TrayAndWait
    if (-not $started) {
        $answer = Read-Host "    $ExeName still isn't running. Press Enter to try again, or type S to skip"
        if ($answer -match '^[sS]') {
            Write-Warn "Skipped. The task stays registered; run this script again once the launch is allowed."
            break
        }
    }
}
if ($started) {
    Write-Success "Running now, and will start automatically at logon."
}

# ─────────────────────────────────────────────
# STEP 5: SHORTCUT ICON
# ─────────────────────────────────────────────

# The exe draws its own "running" icon (folder + green badge) as a multi-size
# .ico, so the shortcuts look exactly like the tray. This runs after the first
# launch above on purpose: by now any antivirus allow has been given. Written to
# a temp file first so a failure can't clobber a good icon from a previous run.
Write-Step "Generating shortcut icon..."

$iconPath = Join-Path $InstallDir $IconFileName
$iconTemp = "$iconPath.tmp"
$exported = $false
foreach ($attempt in 1..2) {
    try {
        $export = Start-Process -FilePath $ExePath -ArgumentList "--export-icon", "`"$iconTemp`"" `
            -WorkingDirectory $InstallDir -Wait -PassThru
        if ($export.ExitCode -eq 0 -and (Test-Path $iconTemp)) { $exported = $true; break }
    } catch {
        # Start-Process throws a terminating error when the launch itself is refused
        # (e.g. "Access is denied" from antivirus); a moment later it may be allowed.
        Write-Warn "Couldn't run $ExeName to draw the icon: $($_.Exception.Message)"
    }
    Start-Sleep -Seconds 1
}
if ($exported) {
    Move-Item -Path $iconTemp -Destination $iconPath -Force
    Write-Success "Created $iconPath"
} else {
    Remove-Item $iconTemp -Force -ErrorAction SilentlyContinue
    if (Test-Path $iconPath) {
        Write-Warn "Keeping the existing icon."
    } else {
        Write-Warn "Shortcuts will use the plain folder icon; re-run this script once $ExeName is allowed to run."
    }
}
$iconLocation = if (Test-Path $iconPath) { "$iconPath,0" } else { "shell32.dll,3" }

# ─────────────────────────────────────────────
# STEP 6: SHORTCUTS (relaunch by hand)
# ─────────────────────────────────────────────

Write-Step "Creating 'Syncthing Monitor' shortcuts..."

Remove-Item $LegacyStartupLnk -Force -ErrorAction SilentlyContinue
$ws = New-Object -ComObject WScript.Shell
foreach ($dir in $ShortcutDirs) {
    $lnk = Join-Path $dir $ShortcutName
    $sc  = $ws.CreateShortcut($lnk)
    $sc.TargetPath       = $ExePath
    $sc.WorkingDirectory = $InstallDir
    $sc.IconLocation     = $iconLocation          # same folder-with-badge icon as the tray
    $sc.Description      = "Shows whether Syncthing is running, in the system tray."
    $sc.Save()
    Write-Success $lnk
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
