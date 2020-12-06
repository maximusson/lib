=MAX_REPLACE(A2, "-", ".")

Function MAX_REPLACE(rngInput, strFind, strReplace)
'DESCRIPTION: search and replaces values within strings. Note: search string is case sensitive

'INPUT:
'(1) rngInput: range object from Excel [rng object]
'(2) strFind: text element that has to be replaced [string]
'(3) strReplace: text element that contains new value [string]

'OUTPUT:
'(1) replaced value

   For Each objCell In rngInput.Cells
      MAX_REPLACE = Replace(objCell.Text, strFind, strReplace)
      Exit Function
   Next

End Function
