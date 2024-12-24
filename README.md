# PSTransparentWindow

A script to change window transparency in Windows

# Example usage:

```ps
 Add-TransparencyRule                                # Click selection, will prompt for transparency
 Add-TransparencyRule -Transparency 70              # Click selection with preset transparency
 Add-TransparencyRule -ProcessName "Notepad"        # Title selection, will prompt for transparency
 Add-TransparencyRule -ProcessName "Notepad" -Transparency 70  # Title selection with preset transparency

```

### Adding startup task

```

# Change to location of VBS, also update the path in the VBS itself along with the path in the startup-trans.bat
$Action = New-ScheduledTaskAction -Execute "wscript.exe" -Argument """C:\Users\tadghh\Documents\startup-shell.vbs"""
$Trigger = New-ScheduledTaskTrigger -AtLogOn
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
$Principal = New-ScheduledTaskPrincipal -UserId (Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty UserName) -RunLevel Highest -LogonType Interactive

Register-ScheduledTask -TaskName "WindowTransparency" -Action $Action -Trigger $Trigger -Settings $Settings -Principal $Principal -Force
```
