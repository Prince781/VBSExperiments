Dim x
Dim y

set oShell = createobject("wscript.shell")
Set Sapi = Wscript.CreateObject("SAPI.SpVoice") 

x = InputBox("Would you like to shutdown your PC after a period of time? Answer must be exactly 'Yes' or 'No'. (Case Sensitive)","AUTOMATIC SYSTEM SHUTDOWN! by Princeton Ferro")

if x = "Yes" then
y = InputBox("Enter time left for shutdown to occur automatically(time in minutes...):","Automatic System Shutdown")
sapi.speak y + "Minutes left for shutdown"
oShell.Run "shutdown.exe -s -t " & (y * 60) & " -f -c ""System is now set for automatic shutdown! You can now go to sleep. Don't wait for this to finish. Good night! See you tomorrow! To terminate, run this and type 'Stop Shutdown'."""

If Err.Number <> 0 Then
y=Msgbox("You did not type a number! Go back and try again.",64,"Error")
End If

end if 
if x = "No" then
x=msgbox("Okay, as you wish good bye, script organized by Princeton Ferro, created by someone else.",64,"Script Aborted")
end if 

if x = "Stop Shutdown" then
oShell.run "shutdown.exe -a"
x=msgbox("Shutdown has been terminated.",64,"Shutdown Aborted")
end if
