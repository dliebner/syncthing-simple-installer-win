' launch-tray.vbs - starts the tray app (SyncthingTray.ps1, in this folder)
' with no visible window at all. The "Syncthing Monitor" shortcuts created by
' Install-SyncthingTray.ps1 point at this file.
' -STA is required by the tray/menu UI; -NonInteractive stops any stray prompt
' from hanging an invisible window.
Dim shell, scriptDir
Set shell = CreateObject("WScript.Shell")
scriptDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
' Run the tray from its own folder, whatever folder this .vbs was launched from,
' so the tray process never keeps some other folder locked as "in use".
shell.CurrentDirectory = scriptDir
shell.Run "powershell.exe -NoProfile -NonInteractive -STA -ExecutionPolicy Bypass -File """ & scriptDir & "SyncthingTray.ps1""", 0, False
