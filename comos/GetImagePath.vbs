Sub OnCreate()

	Set objComos = Document
	strImagePath = GetImagePath(objComos)
	
	If strImagePath <> "" Then
	  Filename = "C:\Users\coadmin1\Desktop\pid.png"
	End If
	
End Sub

Function GetImagePath(objComos)
' function to load background picture. explanation follows
	
	GetImagePath = ""
	
	If objComos Is Nothing Then Exit Function

	'Load Attributes
	Set objAttrImagePath = objComos.spec("Z00T00001.Z00A00001")
	Set objAttrShowImage = objComos.spec("Z00T00001.Z00A00002") 
	
	'Check Attributes
	If objAttrImagePath Is Nothing Then Exit Function
	If objAttrShowImage Is Nothing Then Exit Function
	
	'Load Path
	strImagePath = objAttrImagePath.DisplayValue
	strImagePath = Replace(strImagePath,Chr(34),"")
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	'Check Path
	If fso.FileExists(strImagePath) = false Then Exit Function
	
	'Show Image?
	If objAttrShowImage.Value = 1 Then
		GetImagePath = strImagePath
	End If
	
End Function
