set ws = WScript.CreateObject("WScript.Shell")
txt = inputbox("enter text to be encoded") 
msgbox encode(txt) 
ws.Run "%windir%\notepad"
wscript.sleep 1000 
ws.sendkeys(encode(txt))
      
function encode(s) 
For i = 1 To Len(s) 
newtxt = Mid(s, i, 1) 
newtxt = Chr(Asc(newtxt)+3) 
coded = coded & newtxt 
Next
encode =  coded
End Function 