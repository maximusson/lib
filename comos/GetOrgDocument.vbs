'example:
Set objOrgDoc = GetOrgDocument(a)

Function GetOrgDocument(objDoc)
' DESCRIPTION: return orginial document
	
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
