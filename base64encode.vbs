Set wShell = CreateObject("Wscript.Shell")
Set oShell = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")
strDesktop = wShell.SpecialFolders("Desktop")
Function Base64Encode(sText) 
    Dim oXML, oNode 
 
    Set oXML = CreateObject("Msxml2.DOMDocument.3.0") 
    Set oNode = oXML.CreateElement("base64") 
    oNode.dataType = "bin.base64" 
    oNode.nodeTypedValue =Stream_StringToBinary(sText) 
    Base64Encode = oNode.text 
    Set oNode = Nothing 
    Set oXML = Nothing 
End Function 
 
Function Base64Decode(vCode) 
    Dim oXML, oNode 
 
    Set oXML = CreateObject("Msxml2.DOMDocument.3.0") 
    Set oNode = oXML.CreateElement("base64") 
    oNode.dataType = "bin.base64" 
    oNode.text = vCode 
    Base64Decode = Stream_BinaryToString(oNode.nodeTypedValue) 
    Set oNode = Nothing 
    Set oXML = Nothing 
End Function 
 
'Stream_StringToBinary Function 
'2003 Antonin Foller, http://www.motobit.com 
'Text - string parameter To convert To binary data 
Function Stream_StringToBinary(Text) 
  Const adTypeText = 2 
  Const adTypeBinary = 1 
 
  'Create Stream object 
  Dim BinaryStream 'As New Stream 
  Set BinaryStream = CreateObject("ADODB.Stream") 
 
  'Specify stream type - we want To save text/string data. 
  BinaryStream.Type = adTypeText 
 
  'Specify charset For the source text (unicode) data. 
  BinaryStream.CharSet = "us-ascii" 
 
  'Open the stream And write text/string data To the object 
  BinaryStream.Open 
  BinaryStream.WriteText Text 
 
  'Change stream type To binary 
  BinaryStream.Position = 0 
  BinaryStream.Type = adTypeBinary 
 
  'Ignore first two bytes - sign of 
  BinaryStream.Position = 0 
 
  'Open the stream And get binary data from the object 
  Stream_StringToBinary = BinaryStream.Read 
 
  Set BinaryStream = Nothing 
End Function 
 
'Stream_BinaryToString Function 
'2003 Antonin Foller, http://www.motobit.com 
'Binary - VT_UI1 | VT_ARRAY data To convert To a string

text = Inputbox("Enter text to be encoded.","Base 64 Encoder")
text2 = Base64Encode(text)
a=msgbox(text2,0,"Base 64 Encoded Text")
oShell.MinimizeAll
Set f1 = objFSO.CreateTextFile(strDesktop + "\Base_64 Encoded.txt", True)
f1.Write text2
f1.close