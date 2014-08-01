Set Sapi = CreateObject("SAPI.SpVoice")
Set objFSO = CreateObject("Scripting.FileSystemObject")

file=InputBox("Type the name of the file in this directory that you want to speak.","Speaking Computer Test")
filef= strFolderPath + file

If file="default123" Then
sapi.speak("This is a default message. Here is some extra speech. There will also be some extra speech after this as well. Here is some extra speech. If you can here this, then the test is clearly successful.")
ElseIf objFSO.FileExists(filef) Then
Set objFile = objFSO.OpenTextFile(filef, 1)

Do Until objFile.AtEndOfStream
    message = objFile.ReadAll
Loop
objFile.close

sapi.speak(message)
End If
