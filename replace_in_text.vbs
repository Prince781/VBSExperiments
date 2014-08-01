Const ForReading = 1
Const ForWriting = 2

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(strFolderPath & "Assignment List Information.txt", ForReading)

strText = objFile.ReadAll
objFile.Close
strNewText = Replace(strText, "Assigments", "Assignments")

Set objFile = objFSO.OpenTextFile(strFolderPath & "Assignment List Information.txt", ForWriting)
objFile.WriteLine strNewText
objFile.Close