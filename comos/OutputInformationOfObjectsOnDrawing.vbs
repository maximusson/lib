Set GC = CreateObject("ComosROUtilities.GlobalCastings")

Set objDoc = a
objDoc.Report.open

Set objDocRep = objDoc.Report.ReportDocument

Output "we have " & objDocRep.ItemCount & " elements"

For i = 0 To objDocRep.ItemCount - 1
	
	Output "line "  & i+1
	
	Set objDocRepItem = objDocRep.item(i)
	
	' output layer
	layer = objDocRepItem.layer
	output "layer: " & layer
	
	' output device
	Set objDev = GC.GC_GetComosDevice(objDocRepItem)
	If Not objDev Is Nothing Then
		Output "objDevice: " & objDev.name
	End If
	
	' output color
	Set IGraphAtt = GC.GC_GetIGraphicAttributes(objDocRepItem)
	If Not IGraphAtt Is Nothing Then 
		Output "color: " & IGraphAtt.color
	End If
	
	' output cdevice
	Set objRoDevice = GC.GC_GetIRODevice(objDocRepItem)
	If not objRoDevice is nothing then
		output "cdevice: " & objRoDevice.CDeviceFullname 
	End If
	
	Output ""
	
Next

objDoc.Report.close
