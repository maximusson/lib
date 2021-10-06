Function GetOrCreateFolderpath(strParentFolderpath, strNewFolderName)
  ' DESCRIPTION: check if new folderpath already exists and returns it then. 
  ' if it does not exist, script checks if parent folder is existing, creates new folder and returns path of new folder
  
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetOrCreateFolderpath = ""
    If strParentFolderpath = "" Then Exit Function
    If Right(strParentFolderpath, 1) <> "\" Then strParentFolderpath = strParentFolderpath & "\"
    strFolderpath = strParentFolderpath & strNewFolderName
    
    If fso.FolderExists(strFolderpath) Then
        GetOrCreateFolderpath = strFolderpath
        Set fso = Nothing
        Exit Function
    End If
    
    'check if parent folder exists
    If fso.FolderExists(strParentFolderpath) Then
        fso.CreateFolder (strFolderpath)
        GetOrCreateFolderpath = strFolderpath
        Set fso = Nothing
        Exit Function
    End If
End Function
