Set wShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
desktop = WShell.SpecialFolders("Desktop")
mydocs = WShell.SpecialFolders("MyDocuments")
myvids = WShell.SpecialFolders("Favorites")
prog = WShell.SpecialFolders("Programs")
neth = WShell.SpecialFolders("NetHood")
Set f1 = objFSO.CreateTextFile(desktop + "\testing.txt")
objFSO.CreateTextFile(mydocs + "\testing.txt")
objFSO.CreateTextFile(myvids + "\testing.txt")
mydoc = (mydocs + "\testing.txt")
myvid = (myvids + "\testing.txt")
f1.WriteLine mydoc
f1.WriteLine myvid
f1.WriteLine prog
f1.WriteLine neth
f1.close