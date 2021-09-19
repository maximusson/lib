'example:
Call CreateNewExcelFile("C://temp.xlsx")

Sub CreateNewExcelFile(strExcelPath)
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Application.DisplayAlerts = False
	Set objWorkbook=objExcel.workbooks.add()
	objWorkbook.SaveAs strExcelPath
	objWorkbook.Close
	objExcel.Workbooks.Close
	objExcel.Quit
	Set objExcel = Nothing 
End Sub
