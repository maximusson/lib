' EXAMPLE :
strExcelPath = "C:\test.xlsx"
strSheetName = "New Sheet"

Function CreateNewExcelSheet(strExcelPath, strSheetName)
' DESCRIPTION: opens given excel file and adds sheet - if sheet exists or file does not exist, COMOS cancels script
	
' SCRIPT REVISIONS :
' (1) 19-May-2022 : created

' INPUT :
' (1) strExcelPath: new path for excel file - [string]
' (2) strSheetName: name for new excel sheet - [sheet]

' OUTPUT :
' (1) CreateNewExcelSheet: returns true if script ran completely [boolean]
	
	CreateNewExcelSheet = false
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	If strExcelPath = "" Then Exit Function
	If strSheetName = "" Then Exit Function
	If fso.FileExists(strExcelPath) = false Then Exit Function

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
	
	CreateNewExcelSheet = true
				
End Function
