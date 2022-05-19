' EXAMPLE :
Set objDocument = a
strFilepath = "C:\temp\single-doc.pdf"
bExported = ExportComosReportToPdf(objDocument, strFilepath)

' EXAMPLE 2 :
Set colTemp = Project.Workset.GetTempCollection
colDocuments.add a
colDocuments.add b
strFilepath = "C:\temp\multiple-doc.pdf"
bExported = ExportComosReportToPdf(colDocuments, strFilepath)

Function ExportComosReportToPdf(objDocuments, strFilepath)
' DESCRIPTION : exports COMOS document(s) to filesystem as pdf.
' filepath must include pdf extension.

' SCRIPT REVISIONS :
' (1) 13-Feb-2020 : created
' (2) 19-May-2022: beautify script, example 2 added (also possible with document collections)
	
' INPUT :
' (1) objDocuments: comos document object(s) - can either be document or collection of documents - [comos document object(s)]
' (2) strFilepath: path of exported pdf document - [pdf]

' OUTPUT :
' (1) returns true if script ran completely [boolean]

 	ExportComosReportToPdf = false

 	If objReport.SystemType <> 29 Then Exit Function
 	If IsFilepathUniqueAndFolderpathValid(strFilepath) = false Then Exit Function
 	strExtension = GetFileExtension(strFilepath)
 	If strExtension <> "pdf" Then Exit Function

 	Set objExport = CreateObject("Comos.PDFExport.PDFExport")
 	objExport.DoCreateBookMarks = true
	objExport.DoIntelligentExport = true
 	objExport.DoIntelligentExportDocuments = true
 	objExport.DoIntelligentExportLocation = true
 	objExport.DoIntelligentExportUnit = true
 	objExport.DescriptionText = true
 	objExport.NavigatorText = false
 	objExport.SilentMode = true
	objExport.Export strFilepath, objDocuments, Project.Workset

 	ExportComosReportToPdf = true

End Function


Function IsFilepathUniqueAndFolderpathValid(strFilepath)
' DESCRIPTION: checks if filepath is unique. scripts is cancelled if file already exists.
' checks if folder exists. scripts is cancelled if folder does not exist

	IsFilepathUniqueAndFolderpathValid = false

	Set fso = CreateObject("Scripting.FileSystemObject")

	' Check if file exists, exit function if true
 	If fso.FileExists(strFilepath) Then Exit Function

 	' Check if folder exists, exit function if false
 	arrStr = split(strFilepath, "\")
 	If UBound(arrStr) = 0 Then Exit Function
 	strFolderpath = left(strFilepath,len(strFilepath)-len(arrStr(UBound(arrStr))))
 	If fso.FolderExists(strFolderpath) = false Then Exit Function

 	IsFilepathUniqueAndFolderpathValid = true

End Function


Function GetFileExtension(strFilepath)
' DESCRIPTION: returns file extension from a given filepath

	GetFileExtension = ""

 	Set fso = CreateObject("Scripting.FileSystemObject")

 	arrStr = split(strFilepath, ".")

 	' Check if at least one dot appears in path, exit function if not
 	If UBound(arrStr) = 0 Then Exit Function

 	GetFileExtension = LCase(arrStr(UBound(arrStr)))

End Function
