Const fsoForAppend = 8
do
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

On Error Resume Next
Dim objTextStream

If NOT objFSO.FileExists(strFolderPath + "Testfile.txt") Then
objFSO.CreateTextFile(strFolderPath + "Testfile.txt")
End If

'Open the text file
Set objTextStream = objFSO.OpenTextFile(strFolderPath + "Testfile.txt", fsoForAppend)

'Display the contents of the text file
objTextStream.WriteLine "This textfile was automatically modified by a Visual Basic Script."
objTextStream.Close
If err.number <> 0 Then
objFSO.CreateTextFile(strFolderPath + "Testfile.txt")
End If
loop