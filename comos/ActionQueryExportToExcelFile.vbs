Sub ActionQueryExportToExcelFile(Query, strExcelPath, strSheetName)
' DESCRIPTION: exports visible data from COMOS query to excel file, that already exists

	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(strExcelPath) = false Then Exit Sub
	
	' open excel file
	Set excelApp = CreateObject("Excel.Application")
	Set excelFile = excelApp.Workbooks.Open(strExcelpath)
	Set sheet = excelFile.Sheets(1)
	sheet.name = strSheetName

	' fill in data in excel file
	Set colColumns = Query.BaseQuery.Columns
	intRowCount = Query.RowCount
	intColumnCount = colColumns.count
	
	' fill headers
	For i = 1 To intColumnCount
		If colColumns.item(i).Visible = true Then
			sheet.Cells(1,i).Value = colColumns.item(i).Description
		End If
	Next
	
	' fill data
	For i = 1 To intRowCount
		For j = 1 To intColumnCount
			If colColumns.item(j).Visible = true Then
				sheet.cells(i + 1, j).Value = Query.Cell(i,j).Text
			End If
		Next
	Next
	
	' save and close excel file
	'excelFile.saved = true 'close without saving, without prompt
	excelFile.save
	excelFile.close
	excelApp.quit
	Set excelApp = Nothing

End Sub
