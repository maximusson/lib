' EXAMPLE :
intSystemType = 8
strSystemUID = "A0DUEJB2"
objComos = GetObjectBySystemUID(intSystemType, strSystemUID)

Function GetObjectBySystemUID(intSystemType, strSystemUID)
' DESCRIPTION : returns object via load object by type - UNTESTED

' SCRIPT REVISIONS :
' (1) 19-May-2022 : created

' INPUT :
' (1) intSystemType: type of comos object - [integer]
' (2) strSystemUID: system uid of object - [string]

' OUTPUT :
' (1) GetObjectBySystemUID: returns comos object [comos object]
	
	Set GetObjectBySystemUID = Nothing
	If strSystemUID = "" then Exit Function
	
	Set GetObjectBySystemUID = Project.Workset.LoadObjectByType(intSystemType, strSystemUID)

End Function
