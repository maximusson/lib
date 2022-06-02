' EXAMPLE :
=MAX_Join(A1:A4;"-")

Function MAX_Join(rngInput, strSeparator)
' DESCRIPTION : joins text with separator

' SCRIPT REVISIONS :
' (1) 02-Jul-2022 : created

' INPUT :
' (1) rngInput: range object from Excel [rng object]
' (2) strSeparator: string between [string]

' OUTPUT :
' (1) MAX_Join: concatenated string with separator [string]

	MAX_Join = ""
	If rngInput.Cells.Count = 0 Then Exit Function
    
	For Each objCell In rngInput.Cells
		strText = objCell.Text
 		MAX_Join = MAX_Join & strText & strSeparator
	Next
    
	MAX_Join = Left(MAX_Join, Len(MAX_Join) - Len(strSeparator))
    
End Function
