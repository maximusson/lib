Sub Action(Query, QueryBrowser)
	Set colDocs = CreateObject("Scripting.Dictionary")
	counter = 0
	
	For i = 1 To Query.RowCount
		Set Cell = Query.Cell(i, "Object")
		If Not Cell Is Nothing Then
			Set objDoc = Cell.Object
			If Not objDoc Is Nothing Then
				If objDoc.SystemType = 29 Then
					Set objOrgDoc = objDoc.OrgDocument
					If objOrgDoc Is Nothing Then Set objOrgDoc = objDoc
					If objOrgDoc.DocumentType.Name <> "Reference" then
						counter = counter + 1
						colDocs.Add counter, objDoc
					End If
				End If
			End If
		End If
	Next
	
    Call ActionQueryExportDocumentsAsPdfs(colDocs)
	
End Sub

 
  Sub ActionQueryExportDocumentsAsPdfs(colDocs)
' DESCRIPTION: gets a collection of documents as dictionary. key is incremented integer starting from 1, value is objDoc
' folderpath for pdfs is on user's desktop
	
	' basic checks
	If colDocs Is Nothing Then
		MsgBox "No documents for export found!"
		Exit Sub
	End If
	
	If colDocs.Count = 0 Then
		MsgBox "No documents for export found!"
		Exit Sub
	End If
	
	' get folder for export
	strFolderExportPath = GetFolderExportPath()
	If strFolderExportPath = "" Then Exit Sub
	
	' start progressbar
	Set objProgressbar = CreateObject("ComosXMLContent.Progress")
	objProgressbar.Caption = "Exporting documents as pdf... "
	objProgressbar.Percentage = 1
	objProgressbar.Where = 1
	objProgressbar.StartProcess
	
	' export all documents
	For i = 1 To colDocs.Count
		Set objDoc = colDocs.Item(i)
		If Not objDoc Is Nothing Then	

			strBaseFilename = Replace(objDoc.SystemFullName, "|", "-")
			strExtension = ".pdf"
			strFilepath = strFolderExportPath & "\" & strBaseFilename & " [" & GetCurrentTimestampAsString() & "]" & strExtension
			bExported = ExportComosReportToPdf(objDoc, strFilepath)	
				
		End If
		
		' update progressbar
		If objProgressbar.State = 3 Then Exit Sub
		objProgressbar.Percentage = Round(i / colDocs.Count * 100)	
		
	Next
	
	' end progressbar
	objProgressbar.StopProcess
	
	' open folder on filesystem
	Call OpenFileOrFolderOnFileSystem(strFolderExportPath)
	
End Sub


Function GetFolderExportPath()
'DESCRIPTION: creates export folder "comos-exports" on desktop. creates timestamp folder for export within parent folder
	
	GetFolderExportPath = ""
	
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set objShell = CreateObject("Wscript.Shell")
	
	strDesktopPath = objShell.SpecialFolders("Desktop")
	
	If strDesktopPath = "" Then
		MsgBox "Could not get desktop path. Ask your COMOS Admin for support!"
		Exit Function
	End If
	
	' get or create folder comos-exports
	strFolderComosExportsPath = strDesktopPath & "\comos-exports"
	If FSO.FolderExists(strFolderComosExportsPath) = False Then
		FSO.CreateFolder (strFolderComosExportsPath)	
	End If
	
	' get or create timestamp folder
	strFolderTimestampPath = strFolderComosExportsPath & "\" & GetCurrentTimestampAsString()
	
	If FSO.FolderExists(strFolderTimestampPath) = True Then	
		MsgBox "Export folder already exists. Please execute script again!"
		Exit Function
	End If
	
	FSO.CreateFolder (strFolderTimestampPath)	
	GetFolderExportPath = strFolderTimestampPath
	
End Function

 
Sub OpenFileOrFolderOnFileSystem(strPath)
' DESCRIPTION : opens file or folder on filesystem
	
' SCRIPT REVISIONS :
' 11.12.2017: Created
	
' INPUT :
' strPath - filepath or folderpath (string)

	Set FSO = CreateObject("Scripting.FileSystemObject")
	
	If FSO.FileExists(strPath) Or FSO.FolderExists(strPath) Then
		Set objWScript = CreateObject("WScript.shell")
		strExecute = Chr(34) & strPath & Chr(34) & ",2"
		objWScript.Run strExecute		
	End If
	
End Sub

 
Function GetCurrentTimestampAsString()
' DESCRIPTION: returns string from current timestamp, that looks like that: 20210919-151703
	
	strYear = Year(Now)
	strMonth = Right("00" & Month(Now), 2)
	strDay = Right("00" & Day(Now), 2)
	strHour = Right("00" & Hour(Now), 2)
	strMinute = Right("00" & Minute(Now), 2)
	strSecond = Right("00" & Second(Now), 2)

	GetCurrentTimestampAsString = strYear & strMonth & strDay & "-" & strHour & strMinute & strSecond
	
End Function

 
Function ExportComosReportToPdf(objReport, strFilepath)
' DESCRIPTION : exports COMOS report to filesystem as pdf.
' filepath must include pdf extension.
		
' SCRIPT REVISIONS :
' 1 - 13-Feb-2020 - Created
		
' INPUT :
' (1) objComos: object from comos tree - [comos object]
' (2) strFilepath: path of exported pdf document - [pdf]
		
' OUTPUT :	
' (1) returns true if script ran completely [boolean]
		
	ExportComosReportToPdf = False
	
	If objReport.SystemType <> 29 Then Exit Function
	If IsFilepathUniqueAndFolderpathValid(strFilepath) = False Then Exit Function
	
	strExtension = GetFileExtension(strFilepath)
	If strExtension <> "pdf" Then Exit Function
	
	Set objExport = CreateObject("Comos.PDFExport.PDFExport")
	objExport.DoCreateBookMarks = True
	objExport.DoIntelligentExport = True
	objExport.DoIntelligentExportDocuments = True
	objExport.DoIntelligentExportLocation = True
	objExport.DoIntelligentExportUnit = True
	objExport.DescriptionText = True
	objExport.NavigatorText = False
	objExport.SilentMode = True
	objExport.Export strFilepath, objReport, Project.WorkSet
	
	ExportComosReportToPdf = True
		
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

 
Function GetFileExtension(strFilepath)
' DESCRIPTION: returns file extension from a given filepath
	
	GetFileExtension = ""	
	arrStr = split(strFilepath, ".")
	
	' Check if at least one dot appears in path, exit function if not
	If UBound(arrStr) = 0 Then Exit Function
	
	GetFileExtension = LCase(arrStr(UBound(arrStr)))
		
End Function

 
Function GetOwnerByLevel(ByVal objStart, intLevel)
' DESCRIPTION:
' returns owner of object, that has certain level.
' root object for level counting starts from root object (project)
	
' VERSIONS:
' 1 - 02-Aug-2019 - created
	
' INPUTS:
' (1) objStart - start object for searching [comos object]
' (2) intLevel - level of owner [integer]
	
' OUTPUTS:
' (1) owner - returns owner or nothing [comos object]
	
	Set colTemp = objStart.WorkSet.GetTempCollection
	
	Do While (Not objStart.Owner Is Nothing)
		colTemp.Append (objStart)
		Set objStart = objStart.Owner	
	Loop
	
	If colTemp.Count = 0 Or colTemp.Count < intLevel Then
		Set GetOwnerByLevel = Nothing
	Else
		Set GetOwnerByLevel = colTemp.Item(colTemp.Count + 1 - intLevel)	
	End If
		
End Function

 
