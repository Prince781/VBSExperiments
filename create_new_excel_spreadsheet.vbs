Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.Workbooks.Add
objExcel.Cells(1, 1).Value = "Test value"
objExcel.ActiveWorkbook.SaveAs ("C:\Documents and Settings\PUBLIC\Desktop\TestSheet.xls")
objExcel.Quit
