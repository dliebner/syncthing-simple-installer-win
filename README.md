# Syncthing Simple Installer for Windows

A simple PowerShell script to install and configure [Syncthing](https://syncthing.net/) as a Scheduled Task on Windows.

I created this because I couldn't find another small, easily auditable script that gets Syncthing running reliably as a service.

## Features

* **Always Up-To-Date:** Downloads the latest Syncthing release from [its GitHub repo](https://github.com/syncthing/syncthing).
* **Background Daemon:** Creates a Windows Scheduled Task to run Syncthing on login (on the current user).
* **Auto-Updating Support:** Installs to `%LOCALAPPDATA%\Programs\Syncthing`, which allows Syncthing's built-in self-updater to work seamlessly without requiring Administrator privileges.
* **Firewall Configuration:** Automatically creates the necessary Windows Defender Firewall inbound rules.
* **Auto-Generated Uninstaller:** Generates a custom `Uninstall-Syncthing.ps1` script in your installation folder.

## How to Use

Run `Install-Syncthing.ps1`.

Once the script finishes, Syncthing will be running silently in the background. Open your browser and go to **http://localhost:8384** to access the Syncthing Web GUI and set up your folders.

## Advanced Configuration

If you run the script from a terminal, you can customize the installation using parameters:
* `-InstallDir` (Default: `%LOCALAPPDATA%\Programs\Syncthing`)
* `-GuiPort` (Default: `8384`)
* `-StartupDelay` (Default: `30` seconds)
