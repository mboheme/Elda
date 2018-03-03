Dim courant_dir
courant_dir = WScript.ScriptFullName 
courant_dir=left(courant_dir,InStrRev(courant_dir,"\"))

strZipFile = courant_dir + "Elda.zip"         
outFolder = "c:\" 
    
Set objShell = CreateObject( "Shell.Application" )
Set objSource = objShell.NameSpace(strZipFile).Items()
Set objTarget = objShell.NameSpace(outFolder)
intOptions = 256
objTarget.CopyHere objSource, intOptions

msgbox "Extraction termine"