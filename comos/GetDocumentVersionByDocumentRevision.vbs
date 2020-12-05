'example:
Set objVersion = GetVersionByRevision(a)

Function GetVersionByRevision(objRevision)
' Loop over all attributes
' Checks:
' nestedname of link attribut is Y00T00123.Y00A00825
' owner of revision is owner of version
' count of found versions is equal one

  Set GetVersionByRevision = Nothing
  If objRevision Is Nothing Then Exit Function

  Set colBPs = objRevision.BackPointerSpecificationsWithLinkObject
  If colBPs.count = 0 Then Exit Function

  intCount = 0
  For i = 1 To colBPs.count
    Set objAttrLink = colBPs.item(i)       
    If Not objAttrLink Is Nothing Then
      If objAttrLink.NestedName = "Y00T00123.Y00A00825" Then
        Set objActVersion = objAttrLink.GetSpecOwner
        If Not objActVersion Is Nothing Then
          If objActVersion.Owner Is objRevision.Owner Then
            intCount = intCount + 1
            Set objVersion = objActVersion
          End If
        End If
      End If
    End If
  Next
       
  ' Be sure only 1 version
  If intCount = 1 Then
    Set GetVersionByRevision = objVersion  
  End If

End Function
