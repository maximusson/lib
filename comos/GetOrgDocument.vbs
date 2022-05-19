' EXAMPLE :
Set objOrgDoc = GetOrgDocument(a)

Function GetOrgDocument(objDocument)
' DESCRIPTION: return orginial document

' SCRIPT REVISIONS :
' (1) 12-Feb-2022 : created
' (2) 19-May-2022 : beautify script

' INPUT :
' (1) objDocument: comos document [comos document object]

' OUTPUT :
' (2) GetOrgDocument: comos org document [comos document object]
	
	Set GetOrgDocument = Nothing
	
	' basic checks
	If objDoc Is Nothing Then Exit Function
	If objDoc.SystemType <> 29 Then Exit Function
	
	' get org document
	Set objOrgDoc = objDoc.OrgDocument
	If objOrgDoc Is Nothing Then Set objOrgDoc = objDoc
	
	' check if objOrgDoc could not be found, due to missing reference
	If objOrgDoc.DocumentType.Name = "Reference" Then Exit Function
	
	' return
	Set GetOrgDocument = objOrgDoc
	
End Function
