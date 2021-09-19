'example:
Call CreateExcelFile("C://temp.xlsx")

Sub CreateExcelFile(strExcelPath)
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Application.DisplayAlerts = False
	Set objWorkbook=objExcel.workbooks.add()
	objWorkbook.SaveAs strExcelPath
	objWorkbook.Close
	objExcel.Workbooks.Close
	objExcel.Quit
	Set objExcel = Nothing 
End Sub
