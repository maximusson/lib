' EXAMPLE :
Set objDocument = a
bChecked = CheckDocument(objDocument)

Function CheckDocument(objDocument)
' DESCRIPTION : evaluates document check -- unchecked

' REVISIONS :
' (1) 14-Oct-2023 : created
   
' INPUT :
' (1) objDocument: comos document to be checked [comos document]

' OUTPUT :
' (1) CheckDocument: returns true or false depending if script was succesful [boolean]

   CheckDocument = false

   If objDocument Is Nothing Then Exit Function
   If objDocument.SystemType <> 29 Then Exit Function

   Set objDocument = document
   Set comosRepairCon = CreateObject("ComosRepairCon.RepairCon")
   comosRepairCon.CheckDocument objDocument

   CheckDocument = true

End Function
