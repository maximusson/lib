' EXAMPLE :
bAdd = ConvertRtfFileToTxtFile("C:\input.rtf", "C:\output.rtf")

Function ConvertRtfFileToTxtFile(strSourcePathRtf, strDestinationPathTxt)
' DESCRIPTION : converts an rtf file to an txt file

' SCRIPT REVISIONS :
' (1) 01-Jun-2022 : Created

' INPUT :
' (1) strSourcePathRtf: rtf source file - [string]
' (2) strDestinationPathTxt: text file - [unique]

' OUTPUT :
' (1) ConvertRtfFileToTxtFile: returns true if script ran completely [boolean]

	ConvertRtfFileToTxtFile = false
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.fileExists(strSourcePathRtf) = False Then Exit Function
	If fso.fileExists(strDestinationPathTxt) = True Then Exit Function
	If LCase(Right(strSourcePathRtf, 4)) <> ".rtf" Then Exit Function
    
	Set objWord = CreateObject("Word.Application")
	objWord.Documents.Open(strSourcePathRtf).SaveAs strDestinationPathTxt, 2
	objWord.Quit
	Set objWord = Nothing
	Set fso = Nothing
				
	ConvertRtfFileToTxtFile = true

End Function
