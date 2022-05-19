' EXAMPLE :
Output GetLevelOfObject(a)

Function GetLevelOfObject(ByVal objComos)
' DESCRIPTION : returns level of comos object in tree (first level is 1) - WARNING: script not tested

' SCRIPT REVISIONS :
' (1) 13-Oct-2019 : created
' (2) 19-May-2022 : beautify script
	
' INPUT :
' (1) objComos: comos object from tree [comos system object]

' OUTPUT :
' (1) GetLevelOfObject: returns level [integer]

	GetLevelOfObject = -1
	If objComos Is Nothing Then Exit Function
		
	GetLevelOfObject = 1
	Do While (Not objComos.Owner Is Nothing)
		GetLevelOfObject = GetLevelOfObject + 1
		Set objComos = objComos.Owner
	Loop
	
End Function
