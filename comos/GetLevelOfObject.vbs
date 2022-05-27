' EXAMPLE :
Output GetLevelOfObject(Nothing)

Function GetLevelOfObject(ByVal objComos)
' DESCRIPTION : returns level of comos object in tree (project is level 0, first level is 1)

' SCRIPT REVISIONS :
' (1) 13-Oct-2019 : created
' (2) 19-May-2022 : beautify script
' (3) 25-May-2022 : bug fixing, successfully tested	
	
' INPUT :
' (1) objComos: comos object from tree [comos system object]

' OUTPUT :
' (1) GetLevelOfObject: returns level [integer]

	GetLevelOfObject = -1
	If objComos Is Nothing Then Exit Function
		
	GetLevelOfObject = 0
	Do While (Not objComos.Owner Is Nothing)
		GetLevelOfObject = GetLevelOfObject + 1
		Set objComos = objComos.Owner
	Loop
	
End Function
