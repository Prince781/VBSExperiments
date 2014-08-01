strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colFolders = objWMIService.ExecQuery _
    ("Select * from Win32_Directory where name = 'c:\\Test2'")
For Each objFolder in colFolders
    errResults = objFolder.Uncompress
    Wscript.Echo errResults
Next