     set WshShell = WScript.CreateObject("WScript.Shell")
     strDesktop = WshShell.SpecialFolders("Desktop")
     set oShellLink = WshShell.CreateShortcut(strDesktop & "\Shortcut Script.lnk")
     oShellLink.TargetPath = WScript.ScriptFullName
     oShellLink.WindowStyle = 1
     oShellLink.Hotkey = "CTRL+SHIFT+F"
     oShellLink.IconLocation = "notepad.exe, 0"
     oShellLink.Description = "Shortcut Script"
     oShellLink.WorkingDirectory = strDesktop
     oShellLink.Save