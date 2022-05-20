' EXAMPLE :
Set objComos = a
strPdfPath = "C:\test.pdf"
bCreated = CreateNewPdfObjectAndUploadFile(objComos, strPdfPath)

Function CreateNewPdfObjectAndUploadFile(objComos, strPdfPath)
' DESCRIPTION: Creates new pdf object and import pdf from file system. UNTESTED
' used to upload pdfs to base objects
	
' SCRIPT REVISIONS :
' (1) 20-May-2022 : created
	
' INPUT :
' (1) objComos: cdevice where new attribut is created - [comos system object]
' (2) strPdfPath: path of pdf on filesystem - [string]

' OUTPUT :
' (1) CreateNewPdfObjectAndUploadFile: returns true if script ran completely [boolean]	
	
	Set CreateNewPdfObjectAndUploadFile = false

	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(strPdfPath) = false Then Exit Function
	If LCase(Right(strPdfPath,4) <> ".pdf" Then Exit Function
	
	Set objNewDoc = objComos.OwnDocuments.CreateNew
	objNewDoc.DocumentType = Project.DocumentTypes.Item("AdobePDF")

	strPdfFilename = Split(strPdfPath,"\")(UBound(Split(strPdfPath,"\")))

	objNewDoc.Description = strPdfFilename
	objNewDoc.Mode = 2

	fso.CopyFile strPdfPath, objNewDoc.FullFileName
	
	Set CreateNewPdfObjectAndUploadFile = true

End Function
