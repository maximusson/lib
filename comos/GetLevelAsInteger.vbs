' EXAMPLE :
Output GetLevelAsInteger(a)

Function GetLevelAsInteger(ByVal objComos)
' DESCRIPTION : returns level of comos object in tree (first level is 1)

' SCRIPT REVISIONS :
' (1) 13-Oct-2019 : created
' (2) 19-May-2022 : beautify script
	
' INPUT :
' (1) objComos: comos object from tree [comos system object]

' OUTPUT :
' (1) GetLevelAsInteger: returns level [integer]

	GetLevelAsInteger = -1
	If objComos Is Nothing Then Exit Function
		
	GetLevelAsInteger = 0
	Do While (Not objComos.Owner Is Nothing)
		GetLevelAsInteger = GetLevelAsInteger + 1
		Set objComos = objComos.Owner
	Loop
	
End Function
