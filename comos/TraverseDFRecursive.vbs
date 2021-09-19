'example
Call TraverseDFRecursive(a)

Sub TraverseDFRecursive(objRoot)
' DESCRIPTION: traverse through tree. Depth First. Better use the non-recursive version of this script: TraverseDF
	Output objRoot.systemfullname
	
	Set colDevices = objRoot.Devices
	For i = 1 To colDevices.count
		Set objChild = colDevices.item(i)
		Call TraverseDFRecursive(objChild)
	Next
  
End Sub
