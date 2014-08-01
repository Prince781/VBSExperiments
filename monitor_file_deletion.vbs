Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Wscript.Shell")

do until objFSO.FileExists(strFolderPath + protectedfile)
protectedfile=Inputbox("Enter the name (consisting of file name and file extension) of the file in this current folder that you would like to monitor. Please keep in mind that the size of the file will GREATLY affect the performance of this script, so choose wisely.","Enter Name of File to be Protected")
If protectedfile="" Then
objShell.Popup"Okay, bye!",2, "Goodbye"
wscript.quit
End If
If NOT objFSO.FileExists(strFolderPath + protectedfile) Then
objShell.Popup"You did not enter an actual file name. Try again, or type nothing in the prompt box to quit.",3, "Invalid File Name"
End If
loop

copyoffile = strFolderPath + "Copy of " + protectedfile

do
If NOT objFSO.FileExists(strFolderPath + protectedfile) Then

objShell.Popup"Protected file has been deleted! Your file will now respawn. Please wait...",5, "Regenerating File"

objFSO.CopyFile copyoffile , strFolderPath + protectedfile, true
End If

wscript.sleep 2000

If objFSO.FileExists(strFolderPath + protectedfile) Then
wscript.sleep 5000
On Error Resume Next
objFSO.CopyFile strFolderPath + protectedfile , copyoffile, true
End If

loop