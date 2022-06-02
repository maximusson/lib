' EXAMPLE :
=MAX_Split(A2,".",3)

Function MAX_Split(rngInput, strSeparator, intPosition)
' DESCRIPTION : splits the text of a cell by a separator and returns the x-position

' SCRIPT REVISION :
' (1) 02-Jul-2022 : created

' INPUT :
' (1) rngInput : range object from Excel [rng object]
' (2) strSeparator : string that splits input text [string]
' (3) intPosition : position of returned text [integer]

' OUTPUT :
' (1) MAX_Split: concatenated string with separator [string]

	MAX_Split = ""
   
	If rngInput.Cells.Count <> 1 Then Exit Function
	If intPosition < 1 Then Exit Function

	arrStr = Split(rngInput.Cells.Item(1).Text, strSeparator)
	If UBound(arrStr) + 1 >= intPosition Then
		MAX_SPLIT = arrStr(intPosition - 1)
	End If

End Function
