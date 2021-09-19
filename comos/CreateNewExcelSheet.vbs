Sub CreateNewExcelSheet(strExcelPath, strSheetName)
' DESCRIPTION: opens given excel file and adds sheet

	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(strExcelPath) = false Then Exit Sub

	' open excel file
	Set excelApp = CreateObject("Excel.Application")
	Set excelFile = excelApp.Workbooks.Open(strExcelpath)
	
	' check if sheet already exists with new name
	bSheetExists = false
	For Each sheet In excelFile.sheets
		If sheet.Name = strSheetName Then
			bSheetExists = true
		End If
	Next
	
	' add if sheet does not exist
	If bSheetExists = false Then
		excelFile.Sheets.Add(,excelFile.Sheets(excelFile.Sheets.Count)).Name = strSheetName
		'Set objNewSheet = excelFile.sheets.add
		'objNewSheet.Name = strSheetName
	End If
	
	' save and close excel file
	'excelFile.saved = true 'close without saving, without prompt
	excelFile.save
	excelFile.close
	excelApp.quit
	Set excelApp = Nothing
	
End Sub
