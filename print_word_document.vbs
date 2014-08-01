Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Open("c:\TestDocument.doc")
objWord.Visible = False
objDoc.PrintOut()
objWord.Quit
