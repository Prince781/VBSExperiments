tfile = "C:\gtest\Testing.txt"
tfolder = "C:\gtest"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("Shell.Application")
Set wShell = CreateObject("WScript.Shell")
If NOT objFSO.FolderExists(tfolder) Then
objFSO.CreateFolder("C:\gtest")
End if
If NOT objFSO.FileExists(tfile) Then
objFSO.CreateTextFile(tfile)
End if
wscript.sleep 750
Set objFile = objFSO.GetFile("C:\gtest\Testing.txt")
If objFile.attributes and 4 Then      
objFile.attributes = objFile.attributes - 2
objFile.attributes = objFile.attributes - 4
WScript.Echo("System bit is removed.")
Else      
objFile.attributes = objFile.attributes + 2
objFile.attributes = objFile.attributes + 4
WScript.Echo("System bit is set.")   
End If
wShell.run "C:\gtest"
