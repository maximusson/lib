' EXAMPLE :
Call TraverseDF(a)

Sub TraverseDF(objRoot)
' DESCRIPITON: traverses through tree. Depth-First Algorithm

' SCRIPT REVISIONS :
' (1) 19-Sep-2019 : created
' (2) 19-May-2022 : beautify script
	
' INPUT :
' (1) objRoot: comos object - [comos object]]

' OUTPUT :
' () 
	
	If objRoot Is Nothing Then Exit Sub
	
	Set colQueue = CreateObject("System.Collections.ArrayList")
	colQueue.add objRoot
	
	While colQueue.count > 0
		Set objNode = colQueue.item(0)
		colQueue.RemoveAt 0
		Output objNode.SystemFullName
		
		Set colNodes = objNode.devices
		For i = colNodes.count To 1 Step -1
			Set objChildNode = colNodes.item(i)
			colQueue.insert 0, objChildNode
		Next
	Wend

End Sub
