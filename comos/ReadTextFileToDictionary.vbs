' EXAMPLE :
strTextFilePath = "C:/test.txt"
Set dict = ReadTextFileToDictionary(strTextFilePath)
For i = 1 to dict.Count
	output dict(i)	
Next

Function ReadTextFileToDictionary(strTextFilepath)
' DESCRIPTION : reads textfile and save content to dictionary

' SCRIPT REVISIONS :
' (1) 27-May-2022 : Created

' INPUT :
' (1) strTextFilePath: path of textfile - [string]

' OUTPUT :
	' (1) ReadTextFileToDictionary: returns dictionary with content of textfile [dictionary object]


	Set dict = CreateObject("Scripting.Dictionary")
	Set ReadTextFileToDictionary = dict

	Set fso = CreateObject("Scripting.FileSystemObject")	
	If fso.FileExists(strTextFilePath) = false Then Exit Function
		
	Set objFile = fso.OpenTextFile (strTextFilePath, 1)
	intRow = 0
	Do Until objFile.AtEndOfStream
		strLine = objFile.Readline
		row = row + 1
		dict.Add intRow, strLine
	Loop
	objFile.Close
	
	Set ReadTextFileToDictionary = dict
		
End Function
