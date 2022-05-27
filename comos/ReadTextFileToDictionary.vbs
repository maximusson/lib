Function ReadTextFileToDictionary(strFilepath)
Set fso = CreateObject("Scripting.FileSystemObject")
Set dict = CreateObject("Scripting.Dictionary")
filePath = PARAMS
Set file = fso.OpenTextFile (filePath, 1)
row = 0
Do Until file.AtEndOfStream
line = file.Readline
dict.Add row, line
row = row + 1
Loop
file.Close
Set ReadTextFileToDictionary = dict
End Function
