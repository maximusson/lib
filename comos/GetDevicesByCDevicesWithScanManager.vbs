' EXAMPLE :
Set colCDevices = Project.Workset.GetTempCollection
colCDevices.add b
Output colCDevices.count

Set colDevices = GetDevicesByCDevicesWithScanManager(a, colCDevices)
Output colDevices.count

Function GetDevicesByCDevicesWithScanManager(objStart, colCDevices)
' DESCRIPTION : uses scan manager to get collection of objects under a root node with certain cdevices

' SCRIPT REVISIONS :
' (1) 20-May-2022: renamed function
' (2) 25-May-2022: bug fixing, successfully tested

' INPUT :
' (1) objStart: object from comos tree - [comos object]
' (2) colCDevices: collection of cdevices [comos collection - project.workset.gettempcollection]

' OUTPUT :
' (1) GetDevicesByCDevicesWithScanManager: returns collection of found objects  [collection]
	
	Set ws = Project.Workset
	Set GetDevicesByCDevicesWithScanManager = ws.GetTempCollection
	
	If objStart Is Nothing Then Exit Function
	If colCDevices.count = 0 Then Exit Function
	
	Set scanManager = ws.GetScanManager
	
	scanManager.root = objStart
	scanManager.Recursive = True
	scanManager.IncludeRoot = False
	scanManager.SystemType = 8
    
	For i = 1 To colCDevices.count
		Set objCDevice = colCDevices.item(i)
		If Not objCDevice Is Nothing Then 
			If objCDevice.SystemType = 13 Then
				scanManager.CObjects.Append objCDevice
			End If
		End If
	Next
    
	Set GetDevicesByCDevicesWithScanManager = ws.Scan(scanManager)
       
End Function
