Const WINDOW_HANDLE = 0
Const BIF_EDITBOX = &H10
Const BIF_NONEWFOLDER = &H0200
Const BIF_RETURNONLYFSDIRS = &H1

Set objShell = CreateObject("Shell.Application")
Set wshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

'**Browse For Folder To Be Processed
strPrompt = "Please select the folder where the file is located."
intOptions = BIF_RETURNONLYFSDIRS + BIF_NONEWFOLDER + BIF_EDITBOX
strTargetPath = wshShell.SpecialFolders("MyComputer")
strFolderPath = Browse4Folder(strPrompt, intOptions, strTargetPath)

filename=InputBox("Type the name of the file you want to read, now.","File to be Read")

location= strFolderPath & "\" & filename
'The next feature comes in...

Set objFSO = CreateObject("Scripting.FileSystemObject")
On Error Resume Next
Set objFile = objFSO.OpenTextFile(location, 1)

If objFSO.FileExists(location) then
Do Until objFile.AtEndOfStream
    strCharacters = objFile.Read(1000)
    Wscript.Echo strCharacters
Loop
End If

If Not objFSO.FileExists(location) then
Set f1 = objFSO.CreateTextFile(location & " (Replica).txt")
f1.WriteLine "This message is automatically written into a new text document of the specified location if it doesn't exist."
f1.WriteLine "----------------------------------------------------------------"
f1.WriteLine "Date Created: " & Date()
f1.WriteLine "Time Created: " & Time()
Set objFileNew = objFSO.OpenTextFile(location & " (Replica).txt", 1)
Do Until objFileNew.AtEndOfStream
    strCharacters = objFileNew.Read(1000)
    Msgbox strCharacters,0,"Custom Text File: " & location & " (Replica)"
Loop
End If


'**Browse4Folder Function
Function Browse4Folder(strPrompt, intOptions, strRoot)
	Dim objFolder, objFolderItem

	On Error Resume Next

	Set objFolder = objShell.BrowseForFolder(0, strPrompt, intOptions, strRoot)
  	If (objFolder Is Nothing) Then
  		Wscript.Quit
	End If
  	Set objFolderItem = objFolder.Self
  	Browse4Folder = objFolderItem.Path
  	Set objFolderItem = Nothing
  	Set objFolder = Nothing
End Function
