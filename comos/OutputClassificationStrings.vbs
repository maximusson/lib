'example
Call OutputClassificationStrings(objCDev)

Sub OutputClassificationStrings(objCDev)
  For i = 1 To 4
    Output i & vbTab & objCDev.GetClassification(i)
  Next
End Sub
