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
  SyncthingMonitor.cs          The tray icon itself; compiled on your machine at install time
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

It starts automatically at logon through a scheduled task, `Syncthing Monitor (<username>)`, the same mechanism the installer uses for Syncthing itself. If the icon ever goes missing, reopen **Syncthing Monitor** from the Start menu or the desktop shortcut; a second copy won't be started if one is already running. The shortcuts use the same folder icon, but with a badge that is half green and half red, so it can't be mistaken for the live indicator in the tray; it's generated at install time into `SyncthingMonitor.ico` next to `syncthing.exe`.

There is deliberately no Exit item. If you need to stop it, hold **Shift** while right-clicking the icon to reveal a hidden **Exit** entry.

### How it's built

The tray icon is a small C# program, `tray\SyncthingMonitor.cs`. The installer compiles it on your machine into `SyncthingMonitor.exe` next to `syncthing.exe`, using the C# compiler that is part of the .NET Framework on every Windows 10 and 11. There is no binary in the repo to trust and no build tools to install; you can read the source, and what runs is exactly that. Being a real Windows program rather than a script, it needs no interpreter, no execution-policy bypass, and shows no console window. It's only rebuilt when the source changes.

**Allow-list antivirus (PC Matic and similar):** a freshly compiled program is unknown to these products, so the first launch is blocked until you allow it once. That allow is scoped to this one small exe, not to PowerShell or anything else. No installer can whitelist itself in such a product; that's the point of them. What the installer does instead is start the tray through the same task logon uses, report "Access is denied" when the launch is refused, and then wait for you to allow it and press Enter to retry, so one run is enough. For PC Matic, set SuperShield's *Blocking Notification Method* to *Prompt for Override (Advanced)* **before** installing: the prompt then appears during the install, and **Always Allow** on it is the permanent whitelist. Syncthing itself is unaffected.

### Managing the tray icon on its own

The main installer installs the tray icon by default. To control it separately:

* **Skip it during install:** run `Install-Syncthing.ps1 -NoTray` from a terminal.
* **Add it to an existing install:** run `tray\Install-SyncthingTray.ps1`. It copies the tray files into `%LOCALAPPDATA%\Programs\Syncthing` (next to `syncthing.exe`), builds the exe, creates the task and shortcuts and starts it. Re-running it upgrades a running copy in place.
* **Remove just the tray icon:** run `Install-SyncthingTray.ps1 -Uninstall` (from `tray\` or from the install folder). It stops the tray and removes the task and shortcuts. `Uninstall-Syncthing.ps1` also removes it along with everything else.

Under the hood it polls Syncthing's local REST API every 15 seconds and reads the address and API key from `config.xml`. A few settings (poll interval, notifications, the hidden Exit item) are constants at the top of `tray\SyncthingMonitor.cs`.

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
