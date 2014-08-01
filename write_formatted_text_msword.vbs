Set objWord = CreateObject("Word.Application")
objWord.Visible = False
targetPath = "H:\Memorial Day Essay.doc"
Set objDoc = objWord.Documents.Open(targetPath)
Set objSelection = objWord.Selection

objSelection.Font.Name = "Arial"
objSelection.Font.Size = "18"
objSelection.TypeText "The Ultimate"
objSelection.TypeParagraph()

objSelection.Font.Bold = True
objSelection.Font.Size = "12"
objSelection.TypeText "I think it is important to."
objSelection.TypeParagraph()

'save the Word Document
objWord.Quit

