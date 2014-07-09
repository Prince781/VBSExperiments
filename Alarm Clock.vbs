'Alarm Clock.vbs
Set objFSO = CreateObject("Scripting.FileSystemObject")
thisFolder = replace(WScript.ScriptFullName, WScript.ScriptName, "")
logFile = thisFolder + "alarmInfo.log"
thisTitle = "Alarm Clock Management Wizard"
optList = "newAlarm - Create a new alarm." + vbNewLine + _
		  "deleteAlarm - Delete an existing alarm." + vbNewLine + _
		  "viewAlarms - View all existing alarms and times." + vbNewLine + _
		  "quit - Quit this script." + vbNewLine
strPrompt = msgbox("Hello, and welcome to the Alarm Clock Management Wizard. Continue?", 64 + vbOkCancel, thisTitle)
if (strPrompt = vbCancel) then
msgbox "Created by Princeton Ferro. 2010", 0, thisTitle
WScript.quit
end if

function newAlarm()
properTime = false
do while properTime = false
	timeI = inputbox("Please give the amount of time, in minutes, that you wish the alarm to last for", thisTitle)
	On Error Resume Next
	timeI = eval(timeI * 60000)
	ringTime = eval(timer() + timeI)
		if (err.number <> 0) then
		msgbox "Error! Please enter an integer.", 48, thisTitle
		else
		properTime = true
		end if
loop
nameI = inputbox("Now, give the name of the alarm that you would like to have.", thisTitle)
Set objReg = CreateObject("VBScript.RegExp")
objReg.Pattern = "\|"
objReg.IgnoreCase = True
objReg.Global = True
	if NOT objFSO.FileExists(logFile) then
	Set t = objFSO.CreateTextFile(logFile, true)
	WScript.sleep(2000)
	else
	Set t = objFSO.OpenTextFile(logFile, 8)
	end if
	
	t.WriteLine nameI & "," & ringTime & "|"
	t.Close
end function

optSelectConf = false
do while optSelectConf = false
	optSelect = inputbox("Okay, to start, please select an option from this list:" + vbNewLine + optList, thisTitle)

	optSelectConf = true

	select case optSelect
		case "newAlarm"
			newAlarm()
		case "deleteAlarm"
		case "viewAlarms"
		case "quit"
			quitPrompt = msgbox("Do you wish to quit?", 48 + vbYesNo, thisTitle)
			if (quitPrompt = vbYes) then
			WScript.quit
			else
			optSelectConf = false
			end if
		case else
		optSelectConf = false
	end select
loop
