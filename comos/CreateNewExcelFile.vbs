' EXAMPLE :
bCreated = CreateNewExcelFile("C:\temp.xlsx")

Function CreateNewExcelFile(strExcelPath)
' DESCRIPTION : creates new excel file if file is not existing and folderpath valid

' SCRIPT REVISIONS :
' (1) 19-May-2022 : created

' INPUT :
' (1) strExcelPath: new path for excel file - [string]

' OUTPUT :
' (1) returns true if script ran completely [boolean]
	
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
