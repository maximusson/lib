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
	
	' formation
	sheet.Cells.EntireColumn.AutoFit
	sheet.Cells.EntireRow.AutoFit
	sheet.Rows("1:1").Font.Bold = true
	sheet.Rows("1:1").Font.ThemeColor = 1
	sheet.Rows("1:1").Interior.ThemeColor = 5
	sheet.Rows("1:1").AutoFilter
	sheet.Cells(2,1).Select
	excelApp.ActiveWindow.FreezePanes = true
	
	' save and close excel file
	'excelFile.saved = true 'close without saving, without prompt
	excelFile.save
	excelFile.close
	excelApp.quit
	Set excelApp = Nothing

End Sub
