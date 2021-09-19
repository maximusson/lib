'example
Set objNewTab = CreateNewTab(c, "Z00T00001", "New Tab")

Function CreateNewTab(objCDev, strName, strDescription)
' DESCRIPTION: Creates new tab. 

	Set CreateNewTab = Nothing

	If strName = "" or strDescription = "" Then Exit Function

	Set colSpec = objCDev.OwnSpecifications
	If colSpec.ItemExist(strName) Then Exit Function

	Set objTab = colSpec.CreateNewWithName(strName)
	objTab.Description = strDescription	
	objTab.Unit = "@C"
	
	Set CreateNewTab = objTab

End Function
