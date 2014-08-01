title="Testing Msgbox Functionality"
prompt="This is a test to see how message boxes with custom buttons work."
buttons=vbAbortRetryIgnore


intReturn=MsgBox(prompt,buttons,title)

If intReturn= vbAbort then
msgbox"You clicked abort.",32,"Aborted"
End If

If intReturn= vbRetry then
msgbox"You clicked retry.",32,"Retried"
End If

If intReturn= vbIgnore then
msgbox"You clicked ignore.",32,"Ignored"
End If