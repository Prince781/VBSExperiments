Const ForReading = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    (strDesktop & "Info.txt", ForReading)
Do Until objTextFile.AtEndOfStream
    strNextLine = objTextFile.Readline
    arrTempFilesList = Split(strNextLine , vbcrlf)
    For i = 0 to Ubound(arrTempFilesList)
        Wscript.Echo "File: " & arrTempFilesList(i)
    Next
Loop