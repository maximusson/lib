Function MAX_JOIN(rngInput, strSeparator)
' DESCRIPTION : joins text with separator

' SCRIPT REVISIONS :
' (1) 02-Jul-2022 : created

' INPUT :
' (1) rngInput: range object from Excel [rng object]
' (2) strSeparator: string between [string]

' OUTPUT :
' (1) MAX_JOIN: concatenated string with separator [string]

    MAX_JOIN = ""
    intCount = 0
    For Each objCell In rngInput.Cells
        strText = objCell.Text
        MAX_JOIN = MAX_JOIN & strText & strSeparator
        intCount = intCount + 1
    Next
    
    If intCount = 0 Then Exit Function
    
    MAX_JOIN = Left(MAX_JOIN, Len(MAX_JOIN) - Len(strSeparator))
    
End Function
