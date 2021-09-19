Set GC = CreateObject("ComosROUtilities.GlobalCastings")

Set objDoc = a
Set objDocRep = objDoc.Report.ReportDocument

Output "we have " & objDocRep.ItemCount & " elements"

For i = 0 To objDocRep.ItemCount - 1
	
	Output "line "  & i+1
	
	Set objDocRepItem = objDocRep.item(i)
	layer = objDocRepItem.layer
	
	Set objDev = GC.GC_GetComosDevice(objDocRepItem)
	If Not objDev Is Nothing Then
		Output objDev.name
	End If
	
	Set IGraphAtt = GC.GC_GetIGraphicAttributes(objDocRepItem)
	If Not IGraphAtt Is Nothing Then 
		Output "layer: " & layer & vbTab & "color: " & IGraphAtt.color
	End If
	
	Output ""
	
Next
