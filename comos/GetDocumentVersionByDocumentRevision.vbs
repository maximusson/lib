' EXAMPLE :
Set objRevision = a
Set objVersion = GetVersionByRevision(objRevision)

Function GetVersionByRevision(objRevision)
' DESCRIPTION : tries to return version of document from revision - COMOS iDB
'   Loop over all attributes
'   Checks:
'     nestedname of link attribut is Y00T00123.Y00A00825
'     owner of revision is owner of version
'     count of found versions is equal one
  
' SCRIPT REVISIONS :
' (1) 15-Feb-2020 : Created
' (2) 19-May-2022 : beautify script
	
' INPUT :
' (1) objRevision: revision object - [comos revision document object]

' OUTPUT :
' (1) GetVersionByRevision: returns version of document [comos version document object] 

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
