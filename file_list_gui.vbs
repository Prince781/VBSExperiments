On Error Resume Next

Const WINDOW_HANDLE = 0
Const BIF_EDITBOX = &H10
Const BIF_NONEWFOLDER = &H0200
Const BIF_RETURNONLYFSDIRS = &H1

Set objShell = CreateObject("Shell.Application")
Set wshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

'**Browse For Folder To Be Processed
strPrompt = "Please select the folder to be analyzed, then listed."
intOptions = BIF_RETURNONLYFSDIRS + BIF_NONEWFOLDER + BIF_EDITBOX
strTargetPath = wshShell.SpecialFolders("MyComputer")
strFolderPath = Browse4Folder(strPrompt, intOptions, strTargetPath)

Set objNewFile = objFSO.CreateTextFile(strFolderPath & "\filelist.log", True)
Set objFolder = objFSO.GetFolder(strFolderPath)
Set objColFiles = objFolder.Files
Set objColFolders = objFolder.Folders

For Each file In objColFiles
	objNewFile.WriteLine(file.Name)
Next
For Each folder In objColFolders
	objNewFile.WriteLine(folder.Name)
Next
objNewFile.Close

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
