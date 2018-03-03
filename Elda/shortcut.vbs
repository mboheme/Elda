Set objShell = CreateObject("Wscript.Shell")
strDesktop = objShell.SpecialFolders("Desktop")
Set objShortcut = objShell.CreateShortcut(strDesktop & "\Elda.lnk")
objShortcut.TargetPath = "C:\Elda\Elda.mcr"
objShortcut.Description = "Description."
objShortcut.WorkingDirectory = strDesktop
objShortcut.IconLocation = "C:\Elda\Elda.ico, 0"
objShortcut.Save