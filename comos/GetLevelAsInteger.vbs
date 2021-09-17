'example:
Output GetLevelAsInteger(a)

Function GetLevelAsInteger(objComos)
' DESCRIPTION: returns level of comos object in tree

	GetLevelAsInteger = 0
	If objComos Is Nothing Then Exit Function
	Do While (Not objComos.Owner Is Nothing)
		GetLevelAsInteger = GetLevelAsInteger + 1
		Set objComos = objComos.Owner
	Loop
	
End Function
