' EXAMPLE :
Set colMainObjects = GetCollectionOfQueryMainObjects(TopQuery.MainObject)

Function GetCollectionOfQueryMainObjects(objStart)
' DESCRIPTION:
' using query's main object(s), you never know whether it is a single object or a collection 
' this script converts the topquery.mainobject into a collection that can contain 0, 1, 2 or even more objects
' BE CAREFUL! This script uses error handling. This can be tricky when debugging script.
	
' REVISION:
' 1 : 07-Aug-2019 : created
	
' INPUT: 
' (1) objStart: [comos object] or [collection]
	
' OUTPUT: 
' (2) GetCollectionOfQueryMainObjects: [collection]
	
	Set GetCollectionOfQueryMainObjects = Project.WorkSet.GetTempCollection
	
	If objStart Is Nothing Then Exit Function
	
	On Error Resume Next
	
	intCount = objStart.count ' statement that may result in an error
	If Err.Number > 0 Then
		' on error
		GetCollectionOfQueryMainObjects.add objStart
	Else
		' on no error
		For i = 1 To objStart.count
			GetCollectionOfQueryMainObjects.add objStart.item(i)
		Next
	End If  
	Err.Clear

End Function
