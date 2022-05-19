' EXAMPLE :
strFolderPath = GetOrCreateFolderpath(strParentFolderpath, strNewFolderName)

Function GetOrCreateFolderpath(strParentFolderpath, strNewFolderName)
' DESCRIPTION : check if new folderpath already exists and returns it then. 
' if it does not exist, script checks if parent folder is existing, creates new folder and returns path of new folder

' SCRIPT REVISIONS :
' (1) 11-Dec-2018 : created
' (2) 19-May-2022 : beautify script

' INPUT :
' (1) strParentFolderpath: parent folderpath [string]
' (2) strNewFolderName: name for folder [string]
  
' OUTPUT :
' (1) GetOrCreateFolderpath: path of folder, if folder exists [string]
  
	GetOrCreateFolderpath = ""
	If strParentFolderpath = "" Then Exit Function

	Set fso = CreateObject("Scripting.FileSystemObject")
	
	If Right(strParentFolderpath, 1) <> "\" Then strParentFolderpath = strParentFolderpath & "\"
	strFolderpath = strParentFolderpath & strNewFolderName
    
	' check if folder already exists, if true return path
	If fso.FolderExists(strFolderpath) Then
		GetOrCreateFolderpath = strFolderpath
		Set fso = Nothing
		Exit Function
	End If
    
	' check if parent folder exists, if true, create folder and return path
	If fso.FolderExists(strParentFolderpath) Then
		fso.CreateFolder (strFolderpath)
		GetOrCreateFolderpath = strFolderpath
		Set fso = Nothing
		Exit Function
    End If
			
End Function
