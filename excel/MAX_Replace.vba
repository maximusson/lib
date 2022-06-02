' EXAMPLE :
=MAX_Replace(A2, "-", ".")

Function MAX_Replace(rngInput, strSearchString, strReplaceString)
' DESCRIPTION : replaces a part of a string with another string

' SCRIPT REVISIONS :
' (1) 02-Jul-2022 : created

' INPUT :
' (1) rngInput : range object from Excel [rng object]
' (2) strSearchString : contains search string to be replaced [string]
' (3) strReplaceString : contains replace string [string]

' OUTPUT :
' (1) MAX_Replace : returns modified string [string]

    MAX_Replace = 0
    If rngInput.Cells.Count <> 1 Then Exit Function
    
    MAX_Replace = Replace(rngInput.Cells.Item(1).Text, strSearchString, strReplaceString)
    
End Function
