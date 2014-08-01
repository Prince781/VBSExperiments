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

Set objWord = CreateObject("Word.Application")
objWord.Visible = False
targetPath = strFolderPath & "\File List.doc"
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

objSelection.Font.Name = "Times New Roman"
objSelection.Font.Size = "18"
objSelection.Font.Underline = True
objSelection.TypeText "List of All Files in Directory"
objSelection.TypeParagraph()
objSelection.TypeParagraph()
objSelection.TypeParagraph()

Set objFolder = objFSO.GetFolder(strFolderPath)
Set objColFiles = objFolder.Files

For Each file In objColFiles
objSelection.Font.Underline = False
objSelection.Font.Italic = False
objSelection.Font.Size = "12"
objSelection.Font.Bold = True
objSelection.TypeText(file.Name)
objSelection.TypeParagraph()
Next
For Each folder In objColFiles
objSelection.Font.Underline = False
objSelection.Font.Italic = False
objSelection.Font.Bold = True
objSelection.Font.Size = "12"
objSelection.TypeText(file.Name)
objSelection.TypeParagraph()
Next
objSelection.Font.Underline = False
objSelection.Font.Bold = False
objSelection.Font.Italic = True
objSelection.Font.Size = "12"
objSelection.TypeText "This list was automatically created on " & Date() & "."
'save the Word Document
objDoc.SaveAs strFolderPath & "\File List.doc"
objWord.Quit

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