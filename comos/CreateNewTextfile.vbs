' EXAMPLE :
strTextfilePath = "C:\new.txt"
bCreated = CreateNewTextfile(strTextfilePath)

Function CreateNewTextfile(strTextfilePath)
' DESCRIPTION : creates new textfile if path is valid and

' SCRIPT REVISIONS :
' 1 : 19-May-2022 : Created
' 2 : 19-May-2022 : beautify

' INPUT :
' (1) strTextfilePath: path for new textfile - [string]

' OUTPUT :
' (1) returns true if script ran completely [boolean]
  
	CreateNewTextfile = false
	Set fso = CreateObject("Scripting.FileSystemObject")

	If strTextfilePath = "" Then Exit Function
	If IsFilepathUniqueAndFolderpathValid(strTextfilePath) = false Then Exit Function
     
	Set objFile = fso.CreateTextfile(strTextFilePath)
	Set objFile = nothing
	Set fso = nothing
			
	CreateNewTextfile = true

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
