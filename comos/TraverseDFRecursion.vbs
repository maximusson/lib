' EXAMPLE :
Call TraverseDFRecursion(a)

Sub TraverseDFRecursion(objRoot)
' DESCRIPTION: traverse through tree. Depth First. Better use the non-recursive version of this script: TraverseDF
' NOTE - this script is only for demonstration purpose. DO NOT USE RECURSION, it is ugly 	
	
' SCRIPT REVISIONS :
' (1) 19-Sep-2019 : created
' (2) 19-May-2022 : beautify script
	
' INPUT :
' (1) objRoot: comos object - [comos object]]

' OUTPUT :
' () 
	
	Output objRoot.systemfullname
	
	Set colDevices = objRoot.Devices
	For i = 1 To colDevices.count
		Set objChild = colDevices.item(i)
		Call TraverseDFRecursion(objChild)
	Next
  
End Sub
