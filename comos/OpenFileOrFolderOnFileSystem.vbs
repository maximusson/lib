'example
Call OpenFileOrFolderOnFileSystem("C:\Users\Max\Desktop\spec exporter\export")
'Call OpenFileOrFolderOnFileSystem("C:\Users\Max\Desktop\spec exporter\Book1.xlsx")

Sub OpenFileOrFolderOnFileSystem(strPath)
' DESCRIPTION : opens file or folder on filesystem

' SCRIPT REVISIONS :
' 11.12.2017: Created

' INPUT :
' strPath - filepath or folderpath (string)

	Set fso = CreateObject("Scripting.FileSystemObject")
	
	If fso.FileExists(strPath) or fso.FolderExists(strPath) Then

		Set objWScript = CreateObject("WScript.shell")	
		strExecute = Chr(34) & strPath & Chr(34) & ",2"
		objWScript.Run strExecute 
		
	End If
End Sub
