# SyncthingTray.ps1
# Minimal system-tray indicator for Syncthing ("Syncthing Monitor").
#   Windows folder icon + GREEN badge = Syncthing is running and answering.
#   Windows folder icon + RED badge   = Syncthing is not running / not answering.
# Right-click menu (one item, changes with state):
#   "Sync Now"        when running     -> rescans all folders
#   "Start Syncthing" when not running -> starts it via the scheduled task
# No window. No Exit item for the end user. (Optional hidden Shift+Exit, below.)
#
# Install-Syncthing.ps1 (or tray\Install-SyncthingTray.ps1) copies this file
# next to syncthing.exe and launches it hidden at logon via launch-tray.vbs.
#
# Talks to Syncthing over its local API with curl.exe -k, so a self-signed
# HTTPS cert (if GUI TLS is enabled) is a non-issue.

# --------------------------- CONFIG (edit if needed) ---------------------------
# Folder that contains config.xml (Syncthing's home for this user).
$SyncthingHome = "$env:LOCALAPPDATA\Syncthing"

# Path to syncthing.exe (only used as a fallback to read the API key).
# The installer puts this script next to syncthing.exe; fall back to the default
# install location if it isn't there (e.g. when run straight from the repo).
$SyncthingExe = Join-Path $PSScriptRoot "syncthing.exe"
if (-not (Test-Path $SyncthingExe)) { $SyncthingExe = "$env:LOCALAPPDATA\Programs\Syncthing\syncthing.exe" }

# Scheduled task that starts Syncthing (as created by Install-Syncthing.ps1).
$TaskName = "Syncthing - Start at Logon ($env:USERNAME)"

# How often to check, in seconds.
$PollSeconds = 15

# Pop a small notification when it goes down (what actually reaches the end user).
$NotifyOnDown = $true

# Hidden Exit for admins: hold Shift while right-clicking to reveal an Exit item.
# Invisible to the end user on a normal right-click. Set $false to remove entirely.
$EnableShiftExit = $true

# Named event that Install-SyncthingTray.ps1 / the uninstaller signal to ask a
# running instance to exit cleanly (so its icon is removed, not left as a ghost).
# Must match the name used in those scripts.
$ExitEventName = "SyncthingTrayMonitor_Exit_$env:USERNAME"
# -------------------------------------------------------------------------------

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Single instance: if one is already running for this user, quietly exit. ---
# Makes it safe to re-launch (e.g. from the Start-menu shortcut) if the icon ever
# goes missing: a live instance blocks a duplicate, a dead one gets replaced.
$createdNew = $false
$instanceMutex = New-Object System.Threading.Mutex($true, "SyncthingTrayMonitor_$env:USERNAME", [ref]$createdNew)
if (-not $createdNew) { exit }

# --- Stop request from the installer/uninstaller (see $ExitEventName) ---
$exitEvent = New-Object System.Threading.EventWaitHandle($false, [System.Threading.EventResetMode]::ManualReset, $ExitEventName)
[void]$exitEvent.Reset()   # ignore any stale signal left over from a previous stop

# --- Read address, scheme, and API key from config.xml (best-effort) ---
# Defaults match a fresh Install-Syncthing.ps1 setup (GUI TLS off); config.xml
# wins whenever it can be read.
$GuiAddress = "127.0.0.1:8384"
$Scheme     = "http"
$ApiKey     = ""
try {
    [xml]$cfg = Get-Content -LiteralPath (Join-Path $SyncthingHome "config.xml") -ErrorAction Stop
    if ($cfg.configuration.gui.address) { $GuiAddress = $cfg.configuration.gui.address }
    if ($cfg.configuration.gui.tls)     { $Scheme = if ($cfg.configuration.gui.tls -eq 'true') { 'https' } else { 'http' } }
    if ($cfg.configuration.gui.apikey)  { $ApiKey = $cfg.configuration.gui.apikey }
} catch { }
if (-not $ApiKey -and (Test-Path $SyncthingExe)) {
    try { $ApiKey = (& $SyncthingExe cli config gui apikey get 2>$null).Trim() } catch { }
}
$BaseUrl = "$Scheme`://$GuiAddress"
# If the badge ever reads red while Syncthing is clearly up, the scheme is likely
# wrong for this machine -- set $Scheme = "http" or "https" by hand above.

# --- One-time sanity checks, so a misconfiguration surfaces immediately ---
# (rather than only when someone clicks a button that silently does nothing).
$startupWarnings = @()
if (-not $ApiKey) {
    $startupWarnings += "Couldn't read the Syncthing API key, so 'Sync Now' won't work."
}
if (-not (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue)) {
    $startupWarnings += "Scheduled task '$TaskName' not found, so 'Start Syncthing' won't work."
}

# --- Get the real Windows folder icon from the shell, once ---
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class ShellIcon {
    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Auto)]
    public struct SHFILEINFO {
        public IntPtr hIcon;
        public int    iIcon;
        public uint   dwAttributes;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst=260)] public string szDisplayName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst=80)]  public string szTypeName;
    }
    [DllImport("shell32.dll", CharSet=CharSet.Auto)]
    public static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes, ref SHFILEINFO psfi, uint cbSizeFileInfo, uint uFlags);
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool DestroyIcon(IntPtr hIcon);
}
"@

$script:FolderBitmap = $null
try {
    $info = New-Object "ShellIcon+SHFILEINFO"
    $FILE_ATTRIBUTE_DIRECTORY = 0x10
    $SHGFI_ICON               = 0x100
    $SHGFI_USEFILEATTRIBUTES  = 0x10
    $SHGFI_LARGEICON          = 0x0
    $flags = $SHGFI_ICON -bor $SHGFI_USEFILEATTRIBUTES -bor $SHGFI_LARGEICON
    [void][ShellIcon]::SHGetFileInfo("folder", $FILE_ATTRIBUTE_DIRECTORY, [ref]$info, [System.Runtime.InteropServices.Marshal]::SizeOf($info), $flags)
    if ($info.hIcon -ne [IntPtr]::Zero) {
        $ficon = [System.Drawing.Icon]::FromHandle($info.hIcon)
        $script:FolderBitmap = $ficon.ToBitmap()
        $ficon.Dispose()
        [void][ShellIcon]::DestroyIcon($info.hIcon)
    }
} catch { }

# --- Build a folder icon with a colored status badge in the top-right ---
function New-FolderStatusIcon([System.Drawing.Color]$dot) {
    $size = 32
    $bmp  = New-Object System.Drawing.Bitmap $size, $size
    $g    = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode     = 'AntiAlias'
    $g.InterpolationMode = 'HighQualityBicubic'
    $g.Clear([System.Drawing.Color]::Transparent)

    if ($script:FolderBitmap) {
        $g.DrawImage($script:FolderBitmap, 0, 0, $size, $size)
    } else {
        # Fallback: a simple drawn folder if the shell icon wasn't available.
        $body = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 222, 184, 100))
        $g.FillRectangle($body, 3, 12, 26, 16)
        $g.FillRectangle($body, 3, 8, 12, 6)
        $body.Dispose()
    }

    # Status badge: white halo + colored dot, tucked into the top-right corner.
    $bd     = 14
    $ringW  = 2
    $outerD = $bd + 2 * $ringW
    $cx     = $size - [int]($bd / 2) - 2
    $cy     = [int]($bd / 2) + 2
    [int]$ox = $cx - [int]($outerD / 2)
    [int]$oy = $cy - [int]($outerD / 2)
    [int]$ix = $cx - [int]($bd / 2)
    [int]$iy = $cy - [int]($bd / 2)

    $white = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)
    $g.FillEllipse($white, $ox, $oy, $outerD, $outerD)
    $brush = New-Object System.Drawing.SolidBrush $dot
    $g.FillEllipse($brush, $ix, $iy, $bd, $bd)
    $white.Dispose(); $brush.Dispose(); $g.Dispose()

    return [System.Drawing.Icon]::FromHandle($bmp.GetHicon())
}
$IconUp   = New-FolderStatusIcon ([System.Drawing.Color]::FromArgb(40, 170, 70))
$IconDown = New-FolderStatusIcon ([System.Drawing.Color]::FromArgb(200, 60, 60))

$notify = New-Object System.Windows.Forms.NotifyIcon
$notify.Icon    = $IconDown
$notify.Text    = "Syncthing: checking..."
$notify.Visible = $true

$menu = New-Object System.Windows.Forms.ContextMenuStrip
$notify.ContextMenuStrip = $menu

# State
$script:lastUp      = $null                 # last confirmed up/down
$script:menuState   = $null                 # what the menu currently reflects
$script:missCount   = 0                     # consecutive failed health checks
$script:lastBalloon = [datetime]::MinValue  # for the "stopped" balloon cooldown
$script:fastUntil   = [datetime]::MinValue  # end of any fast-poll burst
$script:normalInterval = [Math]::Max(3, $PollSeconds) * 1000

function Test-Up {
    $out = & curl.exe -k -s -m 3 "$BaseUrl/rest/noauth/health" 2>$null
    return ($LASTEXITCODE -eq 0 -and $out -match '"status"\s*:\s*"OK"')
}

function Start-FastPoll {
    # Poll every second for up to 30s so the badge reacts within ~1s of a change.
    $script:fastUntil = (Get-Date).AddSeconds(30)
    $timer.Interval = 1000
}

function Do-Start {
    & schtasks.exe /Run /TN $TaskName 2>$null | Out-Null
    # schtasks is an external program: it reports failure via exit code, not by
    # throwing, so this is the check that actually catches a bad task name.
    if ($LASTEXITCODE -ne 0) {
        $notify.ShowBalloonTip(4000, "Syncthing", "Couldn't start Syncthing: the scheduled task '$TaskName' wasn't found or couldn't run.", [System.Windows.Forms.ToolTipIcon]::Warning)
        return
    }
    $notify.Text = "Syncthing: starting..."
    Start-FastPoll
}

function Do-SyncNow {
    # -f makes curl fail on HTTP errors (e.g. 403 from a bad API key), so an
    # auth failure isn't reported as a successful sync.
    & curl.exe -k -s -f -m 10 -X POST -H "X-API-Key: $ApiKey" "$BaseUrl/rest/db/scan" 2>$null | Out-Null
    if ($LASTEXITCODE -eq 0) {
        $notify.ShowBalloonTip(2000, "Syncthing", "Sync started.", [System.Windows.Forms.ToolTipIcon]::Info)
    } else {
        $why = if (-not $ApiKey) { "no API key was found." } else { "Syncthing may have stopped." }
        $notify.ShowBalloonTip(3000, "Syncthing", "Couldn't trigger a sync: $why", [System.Windows.Forms.ToolTipIcon]::Warning)
        Start-FastPoll   # confirm red quickly if it really has gone down
    }
}

function Stop-Tray {
    # Clean shutdown: stop polling, remove the icon, leave the message loop.
    $timer.Stop()
    $controlTimer.Stop()
    $notify.Visible = $false
    $notify.Dispose()
    [System.Windows.Forms.Application]::Exit()
}

function Set-Menu([bool]$up) {
    $menu.Items.Clear()
    if ($up) {
        $item = $menu.Items.Add("Sync Now")
        $item.add_Click({ Do-SyncNow })
    } else {
        $item = $menu.Items.Add("Start Syncthing")
        $item.add_Click({ Do-Start })
    }
}

function Update-Status {
    $probe = Test-Up
    if ($probe) { $script:missCount = 0 } else { $script:missCount++ }

    # Debounce: require two consecutive misses before declaring "down". A single
    # blip (e.g. the laptop just woke from sleep) shouldn't trigger a false alarm.
    if ($probe)                                                     { $up = $true }
    elseif ($script:missCount -ge 2 -or $script:lastUp -ne $true)   { $up = $false }
    else                                                            { $up = $true }   # one miss while up: hold for now

    $fast = ((Get-Date) -lt $script:fastUntil)

    # Only touch the tray icon when the state actually changes.
    if ($script:lastUp -ne $up) {
        if ($up) { $notify.Icon = $IconUp } else { $notify.Icon = $IconDown }
    }

    if ($up)       { $notify.Text = "Syncthing: running" }
    elseif ($fast) { $notify.Text = "Syncthing: starting..." }
    else           { $notify.Text = "Syncthing: NOT running" }

    # Rebuild the menu when needed, but never while it's open under the mouse.
    if ($script:menuState -ne $up -and -not $menu.Visible) {
        Set-Menu $up
        $script:menuState = $up
    }

    # Balloon on a confirmed up->down transition, with a cooldown to avoid spam
    # if Syncthing is flapping.
    if ($NotifyOnDown -and $script:lastUp -eq $true -and -not $up) {
        if (((Get-Date) - $script:lastBalloon).TotalSeconds -ge 120) {
            $notify.ShowBalloonTip(5000, "Syncthing stopped", "File sync is not running. Right-click the tray icon and choose 'Start Syncthing'.", [System.Windows.Forms.ToolTipIcon]::Warning)
            $script:lastBalloon = Get-Date
        }
    }

    $script:lastUp = $up

    # Leave any fast-poll burst once it's up, or once the 30s window expires.
    if ($up -or -not $fast) {
        $script:fastUntil = [datetime]::MinValue
        if ($timer.Interval -ne $script:normalInterval) { $timer.Interval = $script:normalInterval }
    }
}

if ($EnableShiftExit) {
    $menu.add_Opening({
        foreach ($e in @($menu.Items | Where-Object { $_.Text -eq "Exit" })) { $menu.Items.Remove($e) }
        if ([System.Windows.Forms.Control]::ModifierKeys -band [System.Windows.Forms.Keys]::Shift) {
            $exit = $menu.Items.Add("Exit")
            $exit.add_Click({ Stop-Tray })
        }
    })
}

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = $script:normalInterval
$timer.add_Tick({ try { Update-Status } catch { } })
$timer.Start()

# Watches for a stop request from the installer/uninstaller (see $ExitEventName).
$controlTimer = New-Object System.Windows.Forms.Timer
$controlTimer.Interval = 500
$controlTimer.add_Tick({ if ($exitEvent.WaitOne(0)) { Stop-Tray } })
$controlTimer.Start()

# First check, guarded so a startup hiccup can't kill the app before the loop runs.
try { Update-Status } catch { }

# Surface any config problems found at startup, once.
if ($startupWarnings.Count -gt 0) {
    $notify.ShowBalloonTip(8000, "Syncthing Monitor: check setup", ($startupWarnings -join " "), [System.Windows.Forms.ToolTipIcon]::Warning)
}

[System.Windows.Forms.Application]::Run((New-Object System.Windows.Forms.ApplicationContext))
