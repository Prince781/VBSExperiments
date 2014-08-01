Dim objNet
Set objNet = CreateObject("WScript.NetWork") 
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim strInfo
strInfo = "User Name is     " & objNet.UserName & vbCRLF & _
          "Computer Name is " & objNet.ComputerName & vbCRLF & _
          "Domain Name is   " & objNet.UserDomain
x=Msgbox (strInfo,0,"Computer Information")

strInfo2 = "User Name is " & objNet.UserName
	

Set objFolder = objFSO.CreateFolder(strDesktop + "User Info")
Set f1 = objFSO.CreateTextFile(strDesktop + "User Info" + "\" + "Information.log", True)

'Display the contents of the text file
f1.WriteLine "Initiated by User: " + objNet.ComputerName + "\" + objNet.UserName
'Close the file and clean up
f1.Close

Set objNet = Nothing
