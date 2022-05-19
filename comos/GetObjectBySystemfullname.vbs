' EXAMPLE :
colSystemType = Project.CDevices
strSystemFullname = "@Y|A10|A20"
objComos = GetObjectBySystemfullname(colSystemType, strSystemFullname)

Function GetObjectBySystemfullname(colSystemType, strSystemFullname)
' DESCRIPTION : returns object via load object by type - UNTESTED

' SCRIPT REVISIONS :
' (1) 19-May-2022 : created

' INPUT :
' (1) colSystemType: type of comos collection, Project.CDevices or Project.Devices - [collection]
' (2) strSystemFullname: systemfullname of object - [string]

' OUTPUT :
' (1) GetObjectBySystemfullname: returns comos object [comos object]
	
	Set GetObjectBySystemfullname = Nothing
	If strSystemFullname = "" then Exit Function
	
	Set GetObjectBySystemfullname = Project.GetHierarchyObjectByName(colSystemType, strSystemFullname)

End Function
