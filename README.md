# Syncthing Simple Installer for Windows

A simple PowerShell script to install and configure [Syncthing](https://syncthing.net/) as a Scheduled Task on Windows, plus an optional tray icon that shows whether it's running.

I created this because I couldn't find another small, easily auditable script that gets Syncthing running reliably as a service.

## Features

* **Always Up-To-Date:** Downloads the latest Syncthing release from [its GitHub repo](https://github.com/syncthing/syncthing).
* **Background Daemon:** Creates a Windows Scheduled Task to run Syncthing on login (on the current user).
* **Auto-Updating Support:** Installs to `%LOCALAPPDATA%\Programs\Syncthing`, which allows Syncthing's built-in self-updater to work seamlessly without requiring Administrator privileges.
* **Firewall Configuration:** Automatically creates the necessary Windows Defender Firewall inbound rules.
* **Tray Icon (optional):** A tiny "Syncthing Monitor" in the system tray: a folder icon with a green badge when Syncthing is running and a red one when it isn't. See [Tray icon](#tray-icon-syncthing-monitor).
* **Auto-Generated Uninstaller:** Generates a custom `Uninstall-Syncthing.ps1` script in your installation folder that removes everything, tray icon included.

## What's in the repo

```
Install-Syncthing.ps1          The installer. Run this.
tray/
  Install-SyncthingTray.ps1    Installs just the tray icon (the main installer calls this for you)
  SyncthingTray.ps1            The tray icon itself
```

## How to Use

1. Download the whole repo (**Code > Download ZIP**, or clone it) and extract it. Keep the `tray` folder next to `Install-Syncthing.ps1`.
2. Run `Install-Syncthing.ps1` (see [Running on Windows 11](#running-on-windows-11) if Windows won't let you).

Once the script finishes, Syncthing will be running silently in the background and the tray icon will be in your system tray. Open your browser and go to **http://localhost:8384** to access the Syncthing Web GUI and set up your folders.

If you only download `Install-Syncthing.ps1` on its own it still works; it just skips the tray icon with a warning.

## Tray icon (Syncthing Monitor)

The tray icon is meant for the person who uses the machine day to day, not the person who set it up. It has no window and no settings, just:

* **Folder icon with a green badge:** Syncthing is running. Right-click for **Sync Now**, which rescans all folders.
* **Folder icon with a red badge:** Syncthing is not running. Right-click for **Start Syncthing**, which starts it via the scheduled task. A notification also pops up when Syncthing stops.
* Hovering shows the current state as a tooltip.

It starts automatically at logon. If the icon ever goes missing, reopen **Syncthing Monitor** from the Start menu or the desktop shortcut (a second copy won't be started if one is already running). The shortcuts use the same folder-with-green-badge icon as the tray; it's generated at install time into `SyncthingMonitor.ico` next to `syncthing.exe`.

There is deliberately no Exit item. If you need to stop it, hold **Shift** while right-clicking the icon to reveal a hidden **Exit** entry.

The shortcuts run `powershell.exe` hidden on `SyncthingTray.ps1` directly, with no `.vbs` or other script host in between. That matters on machines with allow-list antivirus such as PC Matic, which block `wscript.exe` by default and would otherwise stop the tray from starting at logon. If your antivirus still complains, the only thing to allow is `powershell.exe` running `SyncthingTray.ps1` from the Syncthing install folder.

### Managing the tray icon on its own

The main installer installs the tray icon by default. To control it separately:

* **Skip it during install:** run `Install-Syncthing.ps1 -NoTray` from a terminal.
* **Add it to an existing install:** run `tray\Install-SyncthingTray.ps1`. It copies the tray files into `%LOCALAPPDATA%\Programs\Syncthing` (next to `syncthing.exe`), creates the shortcuts and starts it. Re-running it upgrades a running copy in place.
* **Remove just the tray icon:** run `Install-SyncthingTray.ps1 -Uninstall` (from `tray\` or from the install folder). `Uninstall-Syncthing.ps1` also removes it along with everything else.

Under the hood it polls Syncthing's local REST API every 15 seconds using `curl.exe` and reads the address and API key from `config.xml`. A few settings (poll interval, notifications, the hidden Exit item) are variables at the top of `tray\SyncthingTray.ps1`.

## Advanced Configuration

If you run the script from a terminal, you can customize the installation using parameters:
* `-InstallDir` (Default: `%LOCALAPPDATA%\Programs\Syncthing`)
* `-GuiPort` (Default: `8384`)
* `-StartupDelay` (Default: `30` seconds)
* `-NoTray` (skip the tray icon)

## Running on Windows 11
Windows 11 restricts running PowerShell scripts by default. Here are two ways to run:

### Option 1: Right-click method
1. **Right-click** `Install-Syncthing.ps1` and select **Run with PowerShell**.

#### Unblocking the script
If the above doesn't work, you may need to unblock the script.
1. **Right-click** `Install-Syncthing.ps1` > **Properties**.
2. At the bottom, check the **Unblock** box and click **OK**.
3. You can now run the script normally.

Tip: if you downloaded the repo as a ZIP, unblock the ZIP file the same way *before* extracting it and every file inside comes out unblocked.

### Option 2: Run in PowerShell terminal with a One-Time Exception
1. Open **PowerShell** (search for it in the Start menu).
2. `cd` to the directory containing `Install-Syncthing.ps1`
3. Run this command to start the installer with a temporary bypass:
```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Install-Syncthing.ps1
```
