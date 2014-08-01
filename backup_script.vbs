Set wShell = CreateObject("Wscript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("Shell.Application")
Set wNet = CreateObject("WScript.NetWork") 
strDesktop = wShell.SpecialFolders("Desktop")
MyDocuments = wShell.SpecialFolders("MyDocuments")
destinlocation = "C:\Computer Backup Files"
desktopfiles = destinlocation + "\Desktop Files"


If NOT objFSO.FolderExists(destinlocation) Then
objFSO.CreateFolder destinlocation
ElseIf NOT objFSO.FolderExists(desktopfiles) Then
objFSO.CreateFolder desktopfiles
Else
filecup=wShell.popup("Oops! It appears that you already have backup folders! Delete?",10,"Delete", vbYesNo)
If filecup = vbYes Then
On Error Resume Next
objFSO.DeleteFolder destinlocation
wShell.popup "Deleting...",3,"Deleting", 0
objFSO.DeleteFolder desktopfiles
wscript.sleep 5000
objFSO.CreateFolder destinlocation
objFSO.CreateFolder desktopfiles
End If
End If

startprompt = wShell.popup("Welcome to the Computer Information Backup Script! Here is a bit of information about this wizard, and how it easily helps you in the tedious and annoying task of backing up your files to either another drive, folder, or zip file. You will continue automatically if nothing is pressed, but if not, press no. (Pressing Yes will make you continue instantly as well.)",10, "Hello, there!", vbYesNo+64)

function startinfo()
sp=msgbox("Okay, well, here's the information about this script:" & vbcrlf & vbcrlf & vbcrlf & _
"1. This script will copy files from your desktop and My Documents area, and back them up onto a folder." & vbcrlf & vbcrlf & _
"2. Also, a feature that records all of the files and their locations into a text file will be implemented. This way, you will be able to check your progress on the amount of copied files." & vbcrlf & vbcrlf & _
"3. You will have the optimization to copy files from your Desktop folder, My Documents folder, and a few other folders as well. In order to make an extra folder option, you will need to list all of the folder locations that you would like to provide for the copying." & vbcrlf & vbcrlf &_
"4. Then, the prompt will ask you to wait for a certain amount of time for the copying process to finish. Once it is finished, more information will provide you with the ending process." & vbcrlf & vbcrlf & _
"Now that you have been briefed on this script, press 'OK' to continue...",0,"Information")
End Function

If startprompt = vbYes Then
startinfo()
ElseIf startprompt = vbNo Then
msgbox"Okay, then. This script will now shut off.",0,"Goodbye"
wscript.quit
Else
startinfo()
End if

If sp <> vbYes Then
wscript.quit
Else
cont=msgbox("Okay, then. First, if you would like to continue, press okay, or cancel. To start, your system will be analyzed to determine how to go through this procedure.",vbOkCancel,"Continue")
End if

function sysanalyze()
Dim strComputer
strComputer="."
    Dim objWMIService, colItems
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48)
Set syslog = objFSO.CreateTextFile("C:\Computer Backup Files\System Info.log")
linebreak = "******************************************************************************************"
syslog.WriteLine linebreak
syslog.WriteLine "System Information"
syslog.WriteLine "Time Taken On: " & Time()
syslog.WriteLine linebreak
syslog.WriteLine ""
For Each objItem in colItems
	    syslog.WriteLine "Boot Device: " & objItem.BootDevice
	    syslog.WriteLine "Build Number: " & objItem.BuildNumber
	    syslog.WriteLine "Build Type: " & objItem.BuildType
	    syslog.WriteLine "Caption: " & objItem.Caption
	    syslog.WriteLine "Code Set: " & objItem.CodeSet
	    syslog.WriteLine "Country Code: " & objItem.CountryCode
	    syslog.WriteLine "Creation ClassName: " & objItem.CreationClassName
	    syslog.WriteLine "CS Creation ClassName: " & objItem.CSCreationClassName
	    syslog.WriteLine "CSD Version: " & objItem.CSDVersion
	    syslog.WriteLine "CS Name: " & objItem.CSName
	    syslog.WriteLine "Current Time Zone: " & objItem.CurrentTimeZone
	    syslog.WriteLine "Debug: " & objItem.Debug
	    syslog.WriteLine "Description: " & objItem.Description
	    syslog.WriteLine "Distributed: " & objItem.Distributed
	    syslog.WriteLine "Foreground Application Boost: " & objItem.ForegroundApplicationBoost
	    syslog.WriteLine "Free Physical Memory: " & objItem.FreePhysicalMemory
	    syslog.WriteLine "Free Space In Paging Files: " & objItem.FreeSpaceInPagingFiles
	    syslog.WriteLine "Free Virtual Memory: " & objItem.FreeVirtualMemory
	    syslog.WriteLine "Install Date: " & objItem.InstallDate
	    syslog.WriteLine "Last Boot Up Time: " & objItem.LastBootUpTime
	    syslog.WriteLine "Local Date Time: " & objItem.LocalDateTime
	    syslog.WriteLine "Locale: " & objItem.Locale
	    syslog.WriteLine "Manufacturer: " & objItem.Manufacturer
	    syslog.WriteLine "Max Number Of Processes: " & objItem.MaxNumberOfProcesses
	    syslog.WriteLine "Max Process Memory Size: " & objItem.MaxProcessMemorySize
	    syslog.WriteLine "Name: " & objItem.Name
	    syslog.WriteLine "Number Of Licensed Users: " & objItem.NumberOfLicensedUsers
	    syslog.WriteLine "Number Of Processes as of Current Time: " & objItem.NumberOfProcesses
	    syslog.WriteLine "Number Of Users: " & objItem.NumberOfUsers
	    syslog.WriteLine "Organization: " & objItem.Organization
	    syslog.WriteLine "OS Language: " & objItem.OSLanguage
	    syslog.WriteLine "OS ProductSuite: " & objItem.OSProductSuite
	    syslog.WriteLine "OS Type: " & objItem.OSType
	    syslog.WriteLine "Other Type Description: " & objItem.OtherTypeDescription
	    syslog.WriteLine "Plus ProductID: " & objItem.PlusProductID
	    syslog.WriteLine "Plus Version Number: " & objItem.PlusVersionNumber
	    syslog.WriteLine "Primary: " & objItem.Primary
	    syslog.WriteLine "Registered User: " & objItem.RegisteredUser
	    syslog.WriteLine "Serial Number: " & objItem.SerialNumber
	    syslog.WriteLine "Service Pack Major Version: " & objItem.ServicePackMajorVersion
	    syslog.WriteLine "Service Pack Minor Version: " & objItem.ServicePackMinorVersion
	    syslog.WriteLine "Size Stored In Paging Files: " & objItem.SizeStoredInPagingFiles
	    syslog.WriteLine "Status: " & objItem.Status
	    syslog.WriteLine "System Device: " & objItem.SystemDevice
	    syslog.WriteLine "System Directory: " & objItem.SystemDirectory
	    syslog.WriteLine "Total Swap SpaceSize: " & objItem.TotalSwapSpaceSize
	    syslog.WriteLine "Total Virtual MemorySize: " & objItem.TotalVirtualMemorySize
	    syslog.WriteLine "Total Visible Memory Size: " & objItem.TotalVisibleMemorySize
	    syslog.WriteLine "Version: " & objItem.Version
	    syslog.WriteLine "Windows Directory: " & objItem.WindowsDirectory
Next
syslog.WriteLine "Current Date Observed On: " & Date()
syslog.Close

End Function

If cont = vbCancel Then

wscript.quit

Else
sysanalyze()

cont2= wShell.popup("Now, to start, your desktop files will be copied. This task should be relatively short, considering most people do not place large files on their desktop. However, if this is not the case, please give some time for the script to finish. You will be notified upon completion. In the meantime, a log of all of the copied files will be written to and typed up in a blank notepad document. During the time which this script is analyzing, it is IMPORTANT that you DO NOT move out of the window in which typing is going on, otherwise key commands could be entered, and could damage your system. If you have read everything, you may press OK to continue.",10, "Copying", 0)
wShell.popup "Now, copying...",3,"Copying", 0

Set desklog = objFSO.CreateTextFile("C:\Computer Backup Files\Desktop Log.log")
Set deskfolder = objFSO.GetFolder(strDesktop)
Set objdesFiles = deskfolder.Files
Set objdesFolders = deskfolder.SubFolders
lnb = "****************************************"
fzz = "*Desktop Log.log                       *"
fzz2 ="*Here is the status of copied files... *"
desklog.WriteLine lnb
desklog.WriteLine fzz
desklog.WriteLine fzz2
desklog.WriteLine lnb
oShell.MinimizeAll
wShell.Run "Notepad.exe"
wscript.sleep 750
If objFSO.FolderExists("C:\Users\") Then
wShell.AppActivate "Untitled - Notepad"
Else
wShell.AppActivate "Untitled - Notepad"
End If

For Each file In objdesFiles
wShell.SendKeys "...Copying file '" & file.Name & "'..." & vbcrlf
wShell.SendKeys "*File Size: " & file.Size & " bytes.." & vbcrlf
desklog.WriteLine ""
desklog.WriteLine "...Copying file '" & file.Name & "'..."
desklog.WriteLine "..File size is " & file.Size & " bytes.."
desklog.WriteLine "..File type is " & file.Type & " .."
desklog.WriteLine "..Short file name is " & file.ShortName & " .."
desklog.WriteLine "..It was created on " & file.DateCreated & " .."
objFSO.CopyFile file , desktopfiles + "\" 
Next
wShell.SendKeys "***Now, copying folders***" & vbcrlf
desklog.WriteLine "-------------------------------------------------"
desklog.WriteLine "***Now, copying folders***"
For Each folder In objdesFolders
wShell.SendKeys "...Copying folder '" & folder.Name & "'..." & vbcrlf
wShell.SendKeys "*Folder Size: " & folder.Size & " bytes.." & vbcrlf
desklog.WriteLine ""
desklog.WriteLine "...Copying folder '" & folder.Name & "'..."
desklog.WriteLine "..Folder size is " & folder.Size & " bytes.."
desklog.WriteLine "..Folder type is " & folder.Type & " .."
desklog.WriteLine "..It was created on " & folder.DateCreated & " .."
objFSO.CopyFolder folder , desktopfiles + "\" 
Next
wShell.SendKeys "You can find the written log of all of this information at the Desktop Log file..."
wscript.sleep 2000
cont3 = msgbox("Now that that is completed with, we shall go onto the My Documents folder...",vbYesNo,"Continue onto My Documents Folder?")
desklog.close
End if

If cont3 = vbNo Then
wscript.quit
Else
cont4=msgbox("Now, let's get onto the My Documents folder. There will NOT be a LIVE recording of files, because it may take too long, but there will be a text file that will record all of the copied files. Click 'OK' to continue.",vbOkCancel,"Continue")
If cont3 = vbCancel Then
wscript.quit
Else
wShell.popup "Loading. Please wait...",2, "Loading...", 0

objFSO.CreateFolder("C:\Computer Backup Files\My Documents Files")
mydocfiles = "C:\Computer Backup Files\My Documents Files"
Set mydoclog = objFSO.CreateTextFile("C:\Computer Backup Files\My Documents Log.log")
Set mydocfolder = objFSO.GetFolder(MyDocuments)
Set objmydFiles = mydocfolder.Files
Set objmydFolders = mydocfolder.SubFolders

lnbr = "****************************************"
fzzr = "*My Documents Log.log                  *"
fzzr2 ="*Here is the status of copied files... *"
desklog.WriteLine lnbr
desklog.WriteLine fzzr
desklog.WriteLine fzzr2
desklog.WriteLine lnbr

For Each file In objmydFiles
mydoclog.WriteLine ""
mydoclog.WriteLine "...Copying file '" & file.Name & "'..."
mydoclog.WriteLine "..File size is " & file.Size & " bytes.."
mydoclog.WriteLine "..File type is " & file.Type & " .."
mydoclog.WriteLine "..Short file name is " & file.ShortName & " .."
mydoclog.WriteLine "..It was created on " & file.DateCreated & " .."
objFSO.CopyFile file , mydocfiles + "\" 
Next

mydoclog.WriteLine "-------------------------------------------------"
mydoclog.WriteLine "***Now, copying folders***"

For Each folder In objmydFolders
mydoclog.WriteLine ""
mydoclog.WriteLine "...Copying folder '" & folder.Name & "'..."
mydoclog.WriteLine "..Folder size is " & folder.Size & " bytes.."
mydoclog.WriteLine "..Folder type is " & folder.Type & " .."
mydoclog.WriteLine "..It was created on " & folder.DateCreated & " .."
objFSO.CopyFolder folder , mydocfiles + "\" 
Next
msgbox"The copying has successfully been completed! Hooray! Now on to the next part...",0,"Complete"
mydoclog.close
End If
End if

