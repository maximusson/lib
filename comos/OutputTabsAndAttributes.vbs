' example
Call OutputTabsAndAttributes(a)

Sub OutputTabsAndAttributes(objComos)
  Output "Tab Name" & vbTab & "Tab Description" & vbTab & "Attribut Name" & vbTab & "Attribut Description"
  Set colTabs = objComos.Specifications
  For i = 1 To colTabs.count
    Set objTab = colTabs.item(i)
    Set colAttr = objTab.specifications
    For j = 1 To colAttr.count
      Set objAttr = colAttr.item(j)
      Output objTab.name & vbTab & objTab.Description & vbTab & objAttr.name & vbTab & objAttr.Description
    Next
  Next
End Sub
