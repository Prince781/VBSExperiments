Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder("C:\Program Files")
If objFolder.Attributes = objFolder.Attributes AND 2 Then
    objFolder.Attributes = objFolder.Attributes XOR 2 
End If