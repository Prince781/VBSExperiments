strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colFolders = objWMIService. _
    ExecQuery("Select * from Win32_Directory where name = 'c:\\Test1'")
For Each objFolder in colFolders
    Wscript.Echo "Archive: " & objFolder.Archive
    Wscript.Echo "Caption: " & objFolder.Caption
    Wscript.Echo "Compressed: " & objFolder.Compressed
    Wscript.Echo "Compression method: " & objFolder.CompressionMethod
    Wscript.Echo "Creation date: " & objFolder.CreationDate
    Wscript.Echo "Encrypted: " & objFolder.Encrypted
    Wscript.Echo "Encryption method: " & objFolder.EncryptionMethod
    Wscript.Echo "Hidden: " & objFolder.Hidden
    Wscript.Echo "In use count: " & objFolder.InUseCount
    Wscript.Echo "Last accessed: " & objFolder.LastAccessed
    Wscript.Echo "Last modified: " & objFolder.LastModified
    Wscript.Echo "Name: " & objFolder.Name
    Wscript.Echo "Path: " & objFolder.Path
    Wscript.Echo "Readable: " & objFolder.Readable
    Wscript.Echo "System: " & objFolder.System
    Wscript.Echo "Writeable: " & objFolder.Writeable
Next