Option Explicit
Dim strComputer, strProcessToKill, objWMIService, colProcess, objProcess

strComputer = "."
strProcessToKill = "notepad.exe"
Set objWMIService = GetObject("winmgmts:" _ 
   & "{impersonationLevel=impersonate}!\\" _ 
   & strComputer _ 
   & "\root\cimv2") 
Set colProcess = objWMIService.ExecQuery _
   ("Select * from Win32_Process Where Name = '" & strProcessToKill & "'")
For Each objProcess in colProcess
'If you want to echo:
'msgbox "... terminating " & objProcess.Name
   objProcess.Terminate()
Next
