Set objWord = CreateObject("Word.Application")
objWord.Visible = False
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

objSelection.Font.Name = "Times New Roman"
objSelection.Font.Size = "18"
objSelection.TypeText "A Note to the Reader"
objSelection.TypeParagraph()

objSelection.Font.Italic = True
objSelection.Font.Size = "12"
objSelection.TypeText "This is a test document by a Visual Basic Script on " & Date()
objSelection.TypeParagraph()

'save the Word Document
objDoc.SaveAs("C:\Documents and Settings\PUBLIC\Desktop\TestDocument.doc")
objWord.Quit