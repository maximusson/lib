' EXAMPLE :
Set objOwner = a
Set objComos = b
strNewName = "h7"
Set objNewComos = CopyObject(objOwner, objComos, strNewName)
If objNewComos Is Nothing Then Output "Nothing"

Function CopyObject(objOwner, objComos, strNewName)
' DESCRIPTION : creates new object if name is unique, works for devices and cdevices so far

' SCRIPT REVISIONS :
' (1) 27-May-2022 : Created and successfully tested

' INPUT :
' (1) objOwner: comos object of new owner - [comos object]
' (2) objComos: comos object - template for copy - [comos object]
' (3) strNewName: name for new object - [string]

' OUTPUT :
' (1) CopyObject: returns object if copying was successful [comos object]

	Set CopyObject = Nothing
	
	If objOwner Is Nothing Then Exit Function
	If objComos Is Nothing Then Exit Function
	If strNewName = "" Then Exit Function
	
	If objComos.SystemType <> objOwner.SystemType Then Exit Function
	
	Select Case objComos.Systemtype
	Case 8 ' devices
		If objOwner.Devices.ItemExist(strNewName) = true Then Exit Function
		
	Case 13 ' cdevices
		If objOwner.CDevices.ItemExist(strNewName) = true Then Exit Function
		
	Case Else
		Exit Function
	
	End Select

	' copy object
	Set objNew = objComos.CopyAll
	objOwner.Paste2(objNew)
	objNew.Name = strNewName
	objNew.Save
	
	' return
	Set CopyObject = objNew
	
End Function
