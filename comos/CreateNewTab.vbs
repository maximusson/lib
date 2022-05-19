'example
Set objNewTab = CreateNewTab(c, "Z00T00001", "New Tab")

Function CreateNewTab(objCDevice, strTabName, strTabDescription)
' DESCRIPTION: Creates new tab. 

' SCRIPT REVISIONS :
' (1) 19-May-2022 : Created
' (2) 19-May-2022 : beautify script
	
' INPUT :
' (1) objCDevice: cdevice where new attribut is created - [comos system object]
' (2) strTabName: name of new tab - [string]
' (3) strTabDescription: description of new tab - [string]

' OUTPUT :
' (1) CreateNewTab: returns tab if script ran completely [boolean]	
	
	Set CreateNewTab = Nothing

	If strTabName = "" or strTabDescription = "" Then Exit Function

	Set colSpec = objCDevice.OwnSpecifications
	If colSpec.ItemExist(strTabName) Then Exit Function

	Set objTab = colSpec.CreateNewWithName(strTabName)
	objTab.Description = strTabDescription	
	objTab.Unit = "@C"
	
	Set CreateNewTab = objTab

End Function
