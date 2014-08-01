Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.DeleteFolder("C:\Copy of Program Files")
x=msgbox("The specified folder(s), along with contents, were deleted and were not sent to the recycle bin.",0,"Folder Deleted")
