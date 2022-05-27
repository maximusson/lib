Sub Action(Query, QueryBrowser)
	
	' EXAMPLE :
	strExcelPath = "C:\export.xlsx"
	strSheetName = "New sheet"
	bExported = ActionQueryExportToExcelFile(Query, strExcelPath, strSheetName)
	
End Sub

Function ActionQueryExportToExcelFile(Query, strExcelPath, strSheetName)
' DESCRIPTION: exports visible data from COMOS query to a new excel file. If file already exists, COMOS cancels export

' SCRIPT REVISIONS :
' (1) 01-May-2022 : created
' (2) 19-May-2022 : beautify script
' (3) 27-May-2022 : added function CreateNewExcelFile(), script tested and working
	
' INPUT :
' (1) Query: query from action function - [comos query object]
' (2) strExcelPath: path for export file - [string]
' (3) strSheetName: name for new excel sheet - [string]
	
' OUTPUT :
' (1) ActionQueryExportToExcelFile: true if script ran completely [boolean]
	
	ActionQueryExportToExcelFile = false
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	If strExcelPath = "" Then Exit Function
	If strSheetName = "" Then Exit Function
	If fso.FileExists(strExcelPath) = true Then Exit Function
	
	' create excel file
	bCreated = CreateNewExcelFile(strExcelPath)
	If bCreated = false Then Exit Function
	
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

	ActionQueryExportToExcelFile = true	
	
End Function

Function CreateNewExcelFile(strExcelPath)
' DESCRIPTION : creates new excel file if file is not existing and folderpath valid

' SCRIPT REVISIONS :
' (1) 19-May-2022 : created

' INPUT :
' (1) strExcelPath: new path for excel file - [string]

' OUTPUT :
' (1) CreateNewExcelFile: returns true if script ran completely [boolean]
	
	CreateNewExcelFile = false
	
	If strExcelPath = "" Then Exit Function
	If IsFilepathUniqueAndFolderpathValid(strExcelPath) = false Then Exit Function

	Set objExcel = CreateObject("Excel.Application")
	objExcel.Application.DisplayAlerts = False
	Set objWorkbook = objExcel.workbooks.add()
	objWorkbook.SaveAs strExcelPath
	objWorkbook.Close
	objExcel.Workbooks.Close
	objExcel.Quit
	Set objExcel = Nothing 

	CreateNewExcelFile = true				
				
End Function

Function IsFilepathUniqueAndFolderpathValid(strFilepath)
' DESCRIPTION: checks if filepath is unique. scripts is cancelled if file already exists.
' checks if folder exists. scripts is cancelled if folder does not exist
	
	IsFilepathUniqueAndFolderpathValid = False
	
	Set FSO = CreateObject("Scripting.FileSystemObject")
	
	' Check if file exists, exit function if true
	If FSO.FileExists(strFilepath) Then Exit Function
	
	' Check if folder exists, exit function if false
	arrStr = split(strFilepath, "\")
	If UBound(arrStr) = 0 Then Exit Function
	
	strFolderpath = Left(strFilepath, Len(strFilepath) - Len(arrStr(UBound(arrStr))))
	If FSO.FolderExists(strFolderpath) = False Then Exit Function
	
	IsFilepathUniqueAndFolderpathValid = True
		
End Function
