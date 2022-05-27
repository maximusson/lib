' EXAMPLE :
strExcelfilePath = "C:\New Microsoft Excel Worksheet.xlsx"
strSheetName = "Tabelle1"
Set dictExcelSheet = ReadExcelFileToDictionary(strExcelfilePath, strSheetName)
If Not dictExcelSheet Is Nothing Then
	intRows = dictExcelSheet.count
	intColumns = dictExcelSheet(1).count
	Output intRows
	Output intColumns 
	' access values via 
	' strValue = dictExcelSheet(intRow)(intColumn)
End If

Function ReadExcelFileToDictionary(strExcelfilePath, strSheetName)
' Function that creates a 2D dictionary (dict in dict)
' stores all displayed text within excel sheet
	Set ReadExcelFileToDictionary = Nothing
	Set dictRows = CreateObject("Scripting.Dictionary")

	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(strExcelfilePath) = false Then Exit Function
	
	Set excelApp = CreateObject("Excel.Application")
	Set excelFile = excelApp.Workbooks.Open(strExcelfilePath)

	' Check if sheet exists
	bSheetExists = False
	For Each objSheet In excelFile.Sheets
		If strSheetName = objSheet.Name Then
			bSheetExists = True
			Exit For
		End If
	Next
	If bSheetExists = false Then 
		excelApp.Quit
		Set excelApp = Nothing
		Set excelFile = Nothing
		Exit Function
	End If
	
	' Get Sheet
	Set objSheet = excelFile.Sheets(strSheetName)
	
	' Define size
	intRows = objSheet.UsedRange.Rows.Count
	intColumns = CInt(objSheet.UsedRange.Columns.Count)
	
	For i = 1 To intRows
		Set dictColumns = CreateObject("Scripting.Dictionary")
		For j = 1 To intColumns
			dictColumns.add j, objSheet.Cells(i,j).Text
		Next
		dictRows.add i, dictColumns
		Set dictColumns = Nothing
	Next
	
	Set ReadExcelFileToDictionary = dictRows
	
	excelApp.Quit
	Set excelApp = Nothing
	Set excelFile = Nothing

End Function
