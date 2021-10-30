'example:
=MAX_MID(A2, 3, 5)

Function MAX_MID(rngInput, intFrom, intLength)
'DESCRIPTION: implements mid function for spreadsheet. left and right function already exists

'INPUT:
'(1) rngInput: range object from Excel [rng object]
'(2) intFrom: starting position [integer]
'(3) intLength: ending position [integer]

'OUTPUT:
'(1) replaced value

   For Each objCell In rngInput.Cells
      MAX_MID = Mid(objCell.Text, intFrom, intLength)
      Exit Function
   Next

End Function
