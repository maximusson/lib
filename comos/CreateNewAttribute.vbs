'example
Set objNewAttr = CreateNewAttribute(c, "Z00A00001", "New Attribute")

Function CreateNewAttribute(objCDev, strName, strDescription)
' DESCRIPTION: Creates new attribut. 

	Set CreateNewAttribute = Nothing

	If strName = "" or strDescription = "" Then Exit Function

	Set colSpec = objCDev.OwnSpecifications
	If colSpec.ItemExist(strName) Then Exit Function

	Set objAttr = colSpec.CreateNewWithName(strName)
	objAttr.Description = strDescription	
	
	Set CreateNewAttribute = objAttr

End Function
