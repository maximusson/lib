Function MAX_JOIN(rngInput, strSeparator)
'DESCRIPTION: joins text with separator

'INPUT:
'(1) rngInput: range object from Excel [rng object]
'(2) strSeparator: string between [string]
    
    MAX_JOIN = ""
    For Each objCell In rngInput.Cells
        strText = objCell.Text
        MAX_JOIN = MAX_JOIN & strText & strSeparator
    Next
    MAX_JOIN = left(MAX_JOIN, len(MAX_JOIN) - len(strSeparator))

End Function
