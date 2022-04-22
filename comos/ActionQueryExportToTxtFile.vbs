Sub Action(Query, QueryBrowser)
	strExportPath = "C:\export.txt"
	Call ActionQueryExportToTxtFile(Query, strExportPath)
End Sub


Sub ActionQueryExportToTxtFile(Query, strExportPath)
' DESCRIPTION: exports visible data from COMOS query to txt file

	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(strExportPath) = true Then Exit Sub
	If Lcase(Right(strExportPath,4)) <> ".txt" then Exit Sub	

	' open txt file
	Set objFile = fso.CreateTextFile(strExportPath)

	' get query row and column information
	Set colColumns = Query.BaseQuery.Columns
	intRowCount = Query.RowCount
	intColumnCount = colColumns.count
	
	' write header
	strHeader = ""
	For i = 1 To intColumnCount
		If colColumns.item(i).Visible = true Then
			strHeader = strHeader & colColumns.item(i).Description
			if i < intColumnCount then
				strHeader = strHeader & vbTab
			end if
		End If
	Next
	objFile.WriteLine strHeader	

	' fill data
	For i = 1 To intRowCount
		strRow = ""
		For j = 1 To intColumnCount
			If colColumns.item(j).Visible = true Then
				strRow = strRow & Query.Cell(i,j).Text
				if j < intColumnCount then
					strRow = strRow & vbTab
				end if
			End If
		Next
		objFile.WriteLine strRow
	Next
	
	' close file
	objFile.close
	Set objFile = Nothing

End Sub
