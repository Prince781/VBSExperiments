Const fsoForAppend = 8

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Open the text file
Dim objTextStream
Set objTextStream = objFSO.OpenTextFile("E:\Testfile.txt", fsoForAppend)

'Display the contents of the text file
objTextStream.WriteLine("This textfile was automatically modified by a Visual Basic Script.")
'Close the file and clean up
objTextStream.Close
Set objTextStream = Nothing
Set objFSO = Nothing
