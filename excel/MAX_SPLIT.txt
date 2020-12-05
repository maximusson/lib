' Example:
=MAX_SPLIT(A2,".",3)


Function MAX_SPLIT(rngInput, strSeparator, intPosition)
'DESCRIPTION: splits the text of a cell by a separator and returns the x-position

'INPUT:
'(1) rngInput: range object from Excel [rng object]
'(2) strSeparator: string that splits input text [string]
'(3) intPosition: position of returned text [integer]

   MAX_SPLIT = "#N/A"

   If intPosition < 1 Then Exit Function

   For Each objCell In rngInput.Cells
      strText = objCell.Text
      arrStr = Split(strText, strSeparator)
      If UBound(arrStr) + 1 >= intPosition Then
         MAX_SPLIT = arrStr(intPosition - 1)
      End If
      Exit Function
   Next

End Function
