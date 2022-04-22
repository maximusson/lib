Sub Action(Query, QueryBrowser)
	strExportPath = "C:\export.txt"
	Call ActionQueryExportToTxtFileUnicode(Query, strExportPath)
End Sub


Sub ActionQueryExportToTxtFileUnicode(Query, strExportPath)
' DESCRIPTION: exports visible data from COMOS query to txt file

	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(strExportPath) = true Then Exit Sub
	If LCase(Right(strExportPath,4)) <> ".txt" then Exit Sub	

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
	strHeader = strHeader & vbCrLf

	' write body
	strBody = ""
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
		strBody = strBody & strRow
		If i < intRowCount Then
			strBody = strBody & vbCrLf
		End If
	Next
	
	' write data
	strData = strHeader & strBody

	' close file
	Set stream = CreateObject("ADODB.Stream")
	stream.Open
	stream.Type = 2     'text
	stream.Position = 0
	stream.Charset = "utf-8"
	stream.WriteText strData
	stream.SaveToFile strExportPath, 2
	stream.Close

End Sub
