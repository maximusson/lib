' EXAMPLE :
strSystemFullname = b.systemfullname
Set objComos = GetCDeviceBySystemfullname(strSystemFullname)
Output objComos.Systemfullname

Function GetCDeviceBySystemfullname(strSystemFullname)
' DESCRIPTION : returns cdevice by systemfullname

' SCRIPT REVISIONS :
' (1) 19-May-2022 : created
' (2) 27-May-2022 : function renamed, only works for CDevices, successfully tested

' INPUT :
' (1) strSystemFullname: systemfullname of object - [string]

' OUTPUT :
' (1) GetCDeviceBySystemfullname: returns comos object [comos object]
	
	Set GetCDeviceBySystemfullname = Nothing
	If strSystemFullname = "" then Exit Function
	
	Set GetCDeviceBySystemfullname = Project.GetHierarchyObjectByName(Project.CDevices, strSystemFullname)

End Function
