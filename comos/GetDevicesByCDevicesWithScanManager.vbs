' EXAMPLE :
Set colDevices = GetDevicesByCDevicesWithScanManager(a, colCDevices)
Output colDevices.count

Function GetDevicesByCDevicesWithScanManager(objStart, colCDevices)
' DESCRIPTION : uses scan manager to get collection of objects under a root node with certain cdevices - UNTESTED

' SCRIPT REVISIONS :
' (1) 20-May-2022: renamed function
	
' INPUT :
' (1) objStart: object from comos tree - [comos object]
' (2) colCDevices: collection of cdevices [comos collection - project.workset.gettempcollection]

' OUTPUT :
' (1) GetDevicesByCDevicesWithScanManager: returns collection of found objects  [collection]
	
	Set ws = Project.Workset
	Set GetDevicesByCDevicesWithScanManager = ws.GetTempCollection
	
	If objStart Is Nothing Then Exit Function
	If UBound(colCDevices) < 0 Then Exit Function
	
	scanManager.root = objStart
	scanManager.Recursive = True
	scanManager.IncludeRoot = False
	scanManager.SystemType = 8
    
	For i = 0 To UBound(colCDevices)
		Set objCDevice = Project.GetCDeviceBySystemFullname(colCDevices.item(i), 1)
        If Not objCDevice Is Nothing Then scanManager.CObjects.Append objCDevice
    Next
    
    Set ScanDevicesForCDevices = ws.Scan(scanManager)
       
End Function
