Select Case WScript.Arguments.Count
	Case 0
		' Default is local computer if none specified
		strComputer = "."
	Case 1
		Select Case WScript.Arguments(0)
			' "?", "-?" or "/?" invoke online help
			Case "?"
				Syntax
			Case "-?"
				Syntax
			Case "/?"
				Syntax
			Case Else
				strComputer = WScript.Arguments(0)
		End Select
	Case Else
		' More than 1 argument is not allowed
		Syntax
End Select

' Define some constants that can be used in this script;
' logoff = 0 (no forced close of applications) or 5 (forced);
' 5 works OK in Windows 2000, but may result in power off in XP
Const EWX_LOGOFF   = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT   = 2
Const EWX_FORCE    = 4
Const EWX_POWEROFF = 8

' Connect to computer
Set OpSysSet = GetObject("winmgmts:{(Shutdown)}//" & strComputer & "/root/cimv2").ExecQuery("select * from Win32_OperatingSystem where Primary=true")

' Actual logoff
for each OpSys in OpSysSet
	OpSys.Win32Shutdown EWX_LOGOFF
next

' Done
WScript.Quit(0)