'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'******************************************************************************
'The VBScript Glant Library
'Version: 1.0
'Creation Date: Sunday May 5, 2010
'------------------------------------------------------------------------------
'The purpose of this file is as a library that can help minimize VBScript 
'programmatic lines. Let this help the user in a more user-friendly way.
'Functions include ones that are for message boxes, input boxes, base_64
'encoding and decoding, creation of text files, and more. 
'******************************************************************************
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'List of functions and descriptions...
'*newfile(path) - Creates a new blank file in the specified path.
'*newdocumentword(path) - Creates a new Word Document in the 
' specified path.
'*checkfolder(path) - Checks to see if a specified folder is 
' existent.
'*checkfile(path) - Checks to see if a specified file is 
' existent.
'*MinimizeEverything() - Minimizes every current window.
'*ActivateWindow(window) - Shows the specified window based on its
' title in the title bar.
'*SpeakSapi(text) - Uses the SAPI system to make the computer utter
' specified words.
'*EchoTextFile(file) - Reads the specified file and displays the 
' contents in a message box.
'*Censor(Input,WordsToSearchFor,CensoredText) - Determines a 
' specified string to put as "input" and then uses the second
' variable, "WordsToSearchFor", to determine what to specifically
' locate in the string. Then, the "CensoredText" variable is where
' to put the word, phrase, or conjunction of leters that will
' replace all occurances of the specified variable in the string.
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Function newfile(path)
Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.CreateTextFile(path)
End Function

Function newdocumentword(path)
Set objWord = CreateObject("Word.Application")
objWord.Visible = False
Set objDoc = objWord.Documents.Add()
objDoc.SaveAs(path)
objWord.Quit
End Function

Function checkfolder(path)
Set objFSO = CreateObject("Scripting.FileSystemObject")
If NOT objFSO.FolderExists(path) Then
objFSO.CreateFolder(path)
msgbox "Directory '" & path & "' doesn't exist. Folder has been created...",64,"Folder Nonexistent"
Else
msgbox "Directory '" & path & "' is existent.",64,"Folder Existent"
End If
End Function

Function checkfile(path)
Set objFSO = CreateObject("Scripting.FileSystemObject")
If NOT objFSO.FileExists(path) Then
objFSO.CreateTextFile(path)
msgbox "File '" & path & "' doesn't exist. File has been created...",64,"File Nonexistent"
Else
msgbox "File '" & path & "' is existent.",64,"File Existent"
End If
End Function

Function MinimizeEverything()
Set wShell = CreateObject("WScript.Shell")
wShell.MinimizeAll
End Function

Function activatewindow(window)
Set wShell = CreateObject("WScript.Shell")
wShell.AppActivate(window)
End Function

Function SpeakSapi(text)
Set sapi = CreateObject("SAPI.SPVoice")
sapi.speak(text)
End Function

Function echotextfile(file)
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(file) Then
Set objFile = objFSO.OpenTextFile(file, 1)
Do Until objFile.AtEndOfStream
    message = objFile.Read(500)
    msgbox message,0,"File Contents"
Loop
objFile.close
Else
msgbox "Sorry, but the specified file at location '" & file & "' does not exist, so therefore, it could not be read.",64,"File Reading Failed"
End If
End Function

Function censor(Input,WordsToSearchFor,CensoredText)
Set objReg = CreateObject("VBScript.RegExp")
objReg.Pattern = WordsToSearchFor
objReg.IgnoreCase = True
objReg.Global = True
Input = objReg.Replace(Input,CensoredText)
End Function

'******************************************************************************
'End of VBScript Glant Library
'******************************************************************************