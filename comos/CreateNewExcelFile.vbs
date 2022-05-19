'example:
Call CreateNewExcelFile("C:\temp.xlsx")

Sub CreateNewExcelFile(strExcelPath)
' DESCRIPTION: creates new excel file
' To Do: check if file already exists -> abort then
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Application.DisplayAlerts = False
	Set objWorkbook = objExcel.workbooks.add()
	objWorkbook.SaveAs strExcelPath
	objWorkbook.Close
	objExcel.Workbooks.Close
	objExcel.Quit
	Set objExcel = Nothing 
End Sub
