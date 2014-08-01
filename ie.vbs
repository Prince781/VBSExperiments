On Error Resume Next
Set ie = CreateObject("InternetExplorer.Application")
Set shell = CreateObject("Shell.Application")
ie.navigate "http://dell.myway.com"
ie.visible = true
sleeptime = 2 * 60000
'...or 2 minutes
sleeptime2 = 5 * 60000

do until a > vbIgnore
On Error Resume Next
wscript.sleep sleeptime
If err.number <> 0 Then
wscript.quit
End If
ie.visible = false
a=msgbox("A serious error has occured. Please logoff your system immediately.",vbAbortRetryIgnore+16,"Error")
If a = vbAbort Then
msgbox"Faled to abort.",16,"Abort Failed"
ElseIf a = vbRetry Then
msgbox"Retry was successful.",64,"Retry Successful"
ie.visible = true
End If
If err.number <> 0 Then
wscript.quit
End If
wscript.sleep sleeptime2
loop