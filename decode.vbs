set x = WScript.createobject("wscript.shell") 
txt = inputbox("enter text to be decoded") 
msgbox encode(txt) 
x.Run "%windir%\notepad"
wscript.sleep 1000 
x.sendkeys encode(txt)     
function encode(s) 
For i = 1 To Len(s) 
newtxt = Mid(s, i, 1) 
newtxt = Chr(Asc(newtxt)-3) 
coded = coded & newtxt 
Next 
encode = coded 
End Function 