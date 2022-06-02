' EXAMPLE :
=MAX_MID(A2, 3, 5)

Function MAX_MID(rngInput, intStartCharacter, intLength)
' DESCRIPTION : implements mid function for spreadsheet. left and right function already exists

' SCRIPT REVISIONS :
' (1) 02-Jul-2022 : created
   
' INPUT :
'( 1) rngInput : range object from Excel [rng object]
' (2) intFrom : starting position [integer]
' (3) intLength : ending position [integer]

' OUTPUT :
' (1) MAX_MID : replaced value

	For Each objCell In rngInput.Cells
		MAX_MID = Mid(objCell.Text, intStartCharacter, intLength)
		Exit Function
	Next

End Function
