' EXAMPLE :
bOpened = OpenFileOrFolderOnFileSystem("C:\export")
'bOpened = OpenFileOrFolderOnFileSystem("C:\export\Book1.xlsx")

Function OpenFileOrFolderOnFileSystem(strPath)
' DESCRIPTION : opens file or folder on filesystem

' SCRIPT REVISIONS :
' (1) 11-Dec-2017 : created
' (2) 19-May-2022 : beautify script, changed from sub to function

' INPUT :
' (1) strPath: filepath or folderpath (string)

' OUTPUT :
' ()
	
	OpenFileOrFolderOnFileSystem = false
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	If fso.FileExists(strPath) = false Then Exit Function
	If fso.FolderExists(strPath) = false Then Exit Function

	Set objWScript = CreateObject("WScript.shell")	
	strExecute = Chr(34) & strPath & Chr(34) & ",2"
	objWScript.Run strExecute 

	OpenFileOrFolderOnFileSystem = true
	
End Function
