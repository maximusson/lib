Function MAX_CountCharacter(rngInput, strSeparator)
' DESCRIPTION : counts a character in a given string

' SCRIPT REVISIONS :
' (1) 02-Jul-2022 : created

' INPUT :
' (1) rngInput : range object from Excel [rng object]
' (2) strSeparator : separator string [string]

' OUTPUT :
' (1) MAX_CountCharacter : returns count of string [Integer]

    MAX_CountCharacter = 0
    If rngInput.Cells.Count <> 1 Then Exit Function
    
    MAX_CountCharacter = UBound(Split(rngInput.Cells.Item(1).Text, strSeparator))
    
End Function
