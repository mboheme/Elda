Set objShell = CreateObject("Wscript.Shell")
strDesktop = objShell.SpecialFolders("Desktop")
Set objShortcut = objShell.CreateShortcut(strDesktop & "\Listing archimonstres.lnk")
objShortcut.TargetPath = "C:\Elda\Listing archimonstres.csv"
objShortcut.Description = "Description."
objShortcut.WorkingDirectory = strDesktop
objShortcut.Save