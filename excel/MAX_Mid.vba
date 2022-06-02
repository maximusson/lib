' EXAMPLE :
=MAX_Mid(A2, 3, 5)

Function MAX_Mid(rngInput, intStartCharacter, intLength)
' DESCRIPTION : implements mid function for spreadsheet. left and right function already exists

' SCRIPT REVISIONS :
' (1) 02-Jul-2022 : created
   
' INPUT :
'( 1) rngInput : range object from Excel [rng object]
' (2) intFrom : starting position [integer]
' (3) intLength : ending position [integer]

' OUTPUT :
' (1) MAX_Mid : replaced value

	MAX_Mid = ""
	If rngInput.Cells.Count <> 1 Then Exit Function

    MAX_Mid = Mid(rngInput.Cells.Item(1).Text, intStartCharacter, intLength)

End Function
