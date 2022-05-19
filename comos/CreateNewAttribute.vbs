' EXAMPLE :
Set objNewAttribute = CreateNewAttribute(c, "Z00A00001", "New Attribute")

Function CreateNewAttribute(objCDevice, strAttributeName, strAttributeDescription)
' DESCRIPTION : creates new attribute on a base object

' SCRIPT REVISIONS :
' (1) 19-May-2022 : Created
' (2) 19-May-2022 : beautify script
	
' INPUT :
' (1) objCDevice: cdevice where new attribut is created - [comos system object]
' (2) strAttributeName: name of new attribut - [string]
' (3) strAttributeDescription: description of new attribute - [string]

' OUTPUT :
' (1) CreateNewAttribute: returns attribute if script ran completely [boolean] 

	Set CreateNewAttribute = Nothing
	
	If objCDevice.SystemType <> 13 Then Exit Function
	If strAttributeName = "" or strAttributeDescription = "" Then Exit Function

	Set colSpec = objCDevice.OwnSpecifications
	If colSpec.ItemExist(strAttributeName) Then Exit Function

	Set objAttribute = colSpec.CreateNewWithName(strAttributeName)
	objAttribute.Description = strAttributeDescription	
	
	Set CreateNewAttribute = objAttribute

End Function
