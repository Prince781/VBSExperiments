'Computer Backup Script 2
'-----------------------------Sets-----------------------------
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set wShell = CreateObject("WScript.Shell")
Set oShell = CreateObject("Shell.Application")
Set wNet = CreateObject("WScript.NetWork")
'-----------------------------Sets-----------------------------
'-------------------Information and Variables------------------
strDesktop = wShell.SpecialFolders("Desktop")
strFolderPath = replace(WScript.ScriptFullName,WScript.ScriptName,"")
strMyDocuments = wShell.SpecialFolders("MyDocuments")
strInfoDir = "C:\Computer Information for " & wNet.ComputerName
'-------------------Information and Variables------------------
Sub quit()
wShell.Popup "Okay, then. Goodbye!",5,"Quitting",64
wscript.quit
End Sub

startprompt = msgbox("Welcome to the Computer Backup Script! This script will help you in backing up your computer's files. To start, click OK to continue.",64+vbOKCancel,"Computer Backup Script")
If startprompt = vbCancel Then
quit()
End If

If objFSO.FileExists("C:\Desktop Log for " & wNet.UserName & ".doc") Then
objFSO.DeleteFile("C:\Desktop Log for " & wNet.UserName & ".doc")
End If

Set desk = objFSO.GetFolder(strDesktop)
Set objWord = CreateObject("Word.Application")
objWord.Visible = False
Set desklog2 = objWord.Documents.Add()
Set deskSelection = objWord.Selection
With deskSelection
.Font.Name = "Arial"
.Font.Size = "20"
.TypeText "List of All Files in Desktop"
.TypeParagraph()
.TypeParagraph()
.Font.Size = "10"
End With
For each file in desk.Files
count = eval(count+1)
Next
If count > 1 Then
grammar = "are"
grammar2 = "files"
ElseIf count = 1 Then
grammar = "is"
grammar2 = "file"
ElseIf count = 0 Then
grammar = "are"
grammar2 = "files"
End If
With deskSelection
.Font.Italic = True
.TypeText "Below " & grammar & " the " & count & " " & grammar2 & " located on your Desktop, listed in alphabetical order."
.Font.Italic = False
.TypeParagraph()
.TypeParagraph()
End With
For each file in desk.Files
With deskSelection
.Font.Bold = True
.Font.Underline = True
.TypeText file.Name
.TypeParagraph()
.Font.Bold = False
.Font.Underline = False
.TypeText "Type: " & file.Type
.TypeParagraph()
If eval(file.Size/1000) < 1 Then
If file.Size <> 1 Then
.TypeText "Size: " & file.Size & " Bytes"
.TypeParagraph()
ElseIf file.Size = 1 Then
.TypeText "Size: " & file.Size & " Byte"
.TypeParagraph()
End If
Else
.TypeText "Size: " & eval(file.Size/1000) & " KB"
.TypeParagraph()
End If
.TypeText "Date Created: " & file.DateCreated
.TypeParagraph()
.TypeParagraph()
End With
Next
desklog2.SaveAs("C:\Desktop Log for " & wNet.UserName & ".doc")
objWord.Quit
Set objWord = Nothing 
Set deskSelection = Nothing 
Set desklog2 = Nothing 

If objFSO.FolderExists(strInfoDir) Then
kjasd = msgbox("It appears that you already have an information directory. Delete?",vbYesNo,"Computer Backup")
If kjasd = vbYes Then
objFSO.DeleteFolder(strInfoDir)
objFSO.CreateFolder(strInfoDir)
End If
ElseIf NOT objFSO.FolderExists(strInfoDir) Then
objFSO.CreateFolder(strInfoDir)
End If

If objFSO.FileExists(strInfoDir & "\Desktop Log for " & wNet.UserName & ".doc") Then
objFSO.DeleteFile(strInfoDir & "\Desktop Log for " & wNet.UserName & ".doc")
End If

objFSO.MoveFile "C:\Desktop Log for " & wNet.UserName & ".doc", strInfoDir & "\Desktop Log for " & wNet.UserName & ".doc"

start2 = msgbox("A log of your desktop has been taken. Would you like to see it now?",64+vbYesNo,"Computer Backup")
If start2 = vbYes Then
Set objWord = CreateObject("Word.Application")
objWord.Visible = True
objWord.Open strInfoDir & "\Desktop Log for " & wNet.UserName & ".doc"
Set objWord = Nothing
wShell.Popup "Initializing...",3,"",0
End If 

start3 = msgbox("Next, this script will copy files from your Desktop and My Documents folder. Continue?",64+vbYesNo,"Computer Backup")