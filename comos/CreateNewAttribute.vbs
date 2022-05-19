'example
Set objNewAttribute = CreateNewAttribute(c, "Z00A00001", "New Attribute")

Function CreateNewAttribute(objCDev, strName, strDescription)
' DESCRIPTION : creates new attribute on a base object

' SCRIPT REVISIONS :
' 1 - 19-May-2022 - Created

' INPUT :
' (1) objCDev: cdevice where new attribut is created - [comos system object]
' (2) strName: name of new attribut - [string]
' (3) strDescription: description of new attribute - [string]

' OUTPUT :
' (1) returns attribute if script ran completely [boolean] 

	Set CreateNewAttribute = Nothing
	
	If objCDev.SystemType <> 13 Then Exit Function
	If strName = "" or strDescription = "" Then Exit Function

	Set colSpec = objCDev.OwnSpecifications
	If colSpec.ItemExist(strName) Then Exit Function

	Set objAttribute = colSpec.CreateNewWithName(strName)
	objAttribute.Description = strDescription	
	
	Set CreateNewAttribute = objAttribute

End Function
