    ' Start PowerPoint.
    Set ppApp = CreateObject("Powerpoint.Application")
 
    ' Make it visible.
    ppApp.Visible = True
 
    ' Add a new presentation.
    Set ppPres = ppApp.Presentations.Add(msoTrue)
 
    ppApp.Save "C:\test.ppt"
ppApp.Quit
