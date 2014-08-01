Sub ListOSProperties( strComputer )
    Dim objWMIService, colItems

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48)
	Set objNewFile = objFSO.CreateTextFile("C:\Ferro System Information.txt", True)
	For Each objItem in colItems
	    objNewFile.WriteLine "Boot Device: " & objItem.BootDevice
	    objNewFile.WriteLine "Build Number: " & objItem.BuildNumber
	    objNewFile.WriteLine "Build Type: " & objItem.BuildType
	    objNewFile.WriteLine "Caption: " & objItem.Caption
	    objNewFile.WriteLine "Code Set: " & objItem.CodeSet
	    objNewFile.WriteLine "Country Code: " & objItem.CountryCode
	    objNewFile.WriteLine "Creation ClassName: " & objItem.CreationClassName
	    objNewFile.WriteLine "CS Creation ClassName: " & objItem.CSCreationClassName
	    objNewFile.WriteLine "CSD Version: " & objItem.CSDVersion
	    objNewFile.WriteLine "CS Name: " & objItem.CSName
	    objNewFile.WriteLine "Current Time Zone: " & objItem.CurrentTimeZone
	    objNewFile.WriteLine "Debug: " & objItem.Debug
	    objNewFile.WriteLine "Description: " & objItem.Description
	    objNewFile.WriteLine "Distributed: " & objItem.Distributed
	    objNewFile.WriteLine "Foreground Application Boost: " & objItem.ForegroundApplicationBoost
	    objNewFile.WriteLine "Free Physical Memory: " & objItem.FreePhysicalMemory
	    objNewFile.WriteLine "Free Space In Paging Files: " & objItem.FreeSpaceInPagingFiles
	    objNewFile.WriteLine "Free Virtual Memory: " & objItem.FreeVirtualMemory
	    objNewFile.WriteLine "Install Date: " & objItem.InstallDate
	    objNewFile.WriteLine "Last Boot Up Time: " & objItem.LastBootUpTime
	    objNewFile.WriteLine "Local Date Time: " & objItem.LocalDateTime
	    objNewFile.WriteLine "Locale: " & objItem.Locale
	    objNewFile.WriteLine "Manufacturer: " & objItem.Manufacturer
	    objNewFile.WriteLine "Max Number Of Processes: " & objItem.MaxNumberOfProcesses
	    objNewFile.WriteLine "Max Process Memory Size: " & objItem.MaxProcessMemorySize
	    objNewFile.WriteLine "Name: " & objItem.Name
	    objNewFile.WriteLine "Number Of Licensed Users: " & objItem.NumberOfLicensedUsers
	    objNewFile.WriteLine "Number Of Processes as of Current Time: " & objItem.NumberOfProcesses
	    objNewFile.WriteLine "Number Of Users: " & objItem.NumberOfUsers
	    objNewFile.WriteLine "Organization: " & objItem.Organization
	    objNewFile.WriteLine "OS Language: " & objItem.OSLanguage
	    objNewFile.WriteLine "OS ProductSuite: " & objItem.OSProductSuite
	    objNewFile.WriteLine "OS Type: " & objItem.OSType
	    objNewFile.WriteLine "Other Type Description: " & objItem.OtherTypeDescription
	    objNewFile.WriteLine "Plus ProductID: " & objItem.PlusProductID
	    objNewFile.WriteLine "Plus Version Number: " & objItem.PlusVersionNumber
	    objNewFile.WriteLine "Primary: " & objItem.Primary
If objFSO.FolderExists("C:\Documents and Settings") Then
	    objNewFile.WriteLine "Quantum Length: " & objItem.QuantumLength
	    objNewFile.WriteLine "Quantum Type: " & objItem.QuantumType
Else
	    objNewFile.Writeline "(Quantum Length and Type not supported on operating systems after Windows XP...)"
End If
	    objNewFile.WriteLine "Registered User: " & objItem.RegisteredUser
	    objNewFile.WriteLine "Serial Number: " & objItem.SerialNumber
	    objNewFile.WriteLine "Service Pack Major Version: " & objItem.ServicePackMajorVersion
	    objNewFile.WriteLine "Service Pack Minor Version: " & objItem.ServicePackMinorVersion
	    objNewFile.WriteLine "Size Stored In Paging Files: " & objItem.SizeStoredInPagingFiles
	    objNewFile.WriteLine "Status: " & objItem.Status
	    objNewFile.WriteLine "System Device: " & objItem.SystemDevice
	    objNewFile.WriteLine "System Directory: " & objItem.SystemDirectory
	    objNewFile.WriteLine "Total Swap SpaceSize: " & objItem.TotalSwapSpaceSize
	    objNewFile.WriteLine "Total Virtual MemorySize: " & objItem.TotalVirtualMemorySize
	    objNewFile.WriteLine "Total Visible Memory Size: " & objItem.TotalVisibleMemorySize
	    objNewFile.WriteLine "Version: " & objItem.Version
	    objNewFile.WriteLine "Windows Directory: " & objItem.WindowsDirectory
	Next
objNewFile.WriteLine "Current Date Observed On: " & Date()
objNewFile.WriteLine "Current Time Observed On: " & Time()
objNewFile.Close

End Sub


' ****************************************************************************
' Main
' ****************************************************************************
Dim strComputer

strComputer="."

ListOSProperties( strComputer )