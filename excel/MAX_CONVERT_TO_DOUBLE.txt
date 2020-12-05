' Example:
=MAX_CONVERT_TO_DOUBLE(A2)

Function MAX_CONVERT_TO_DOUBLE(rngInput)
'DESCRIPTION: trys to convert a cell value into a number. Commas are treated as dots

'INPUT:
'(1) rngInput: range object from Excel [rng object]

'OUTPUT:
'(1) returned value

   For Each objCell In rngInput.Cells
      strText = objCell.Text
      If IsNumeric(strText) Then
         strText = Replace(strText, ",", ".")
         MAX_CONVERT_TO_DOUBLE = CDbl(strText)
      Else
         MAX_CONVERT_TO_DOUBLE = strText
      End If
      Exit Function
   Next

End Function
