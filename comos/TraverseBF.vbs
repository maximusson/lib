' EXAMPLE :
Call TraverseBF(a)

Sub TraverseBF(objRoot)
' DESCRIPTION: traveres through tree. Breadth-First Algorithm
	
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
		For i = 1 To colNodes.count
			Set objChildNode = colNodes.item(i)
			colQueue.add objChildNode
		Next
	Wend
	
End Sub
