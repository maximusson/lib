'Example Code
Set objComos = GetOwnerByLevel(a, 3)

If Not objComos Is Nothing Then
   Output objComos.systemfullname
Else
   Output "nothing"
End If


Function GetOwnerByLevel(ByVal objStart, intLevel)
' DESCRIPTION:
' returns owner of object, that has certain level.
' root object for level counting starts from root object (project)

' VERSIONS:
' 1 - 02-Aug-2019 - created

' INPUTS:
' (1) objStart - start object for searching [comos object]
' (2) intLevel - level of owner [integer]

' OUTPUS:
' (1) owner - returns owner or nothing [comos object]

   Set colTemp = objStart.Workset.GetTempCollection

   Do While (Not objStart.Owner Is Nothing)
   colTemp.Append(objStart)
   Set objStart = objStart.Owner
   Loop

   If colTemp.count = 0 or colTemp.count < intLevel Then
      Set GetOwnerByLevel = Nothing
   Else
      Set GetOwnerByLevel = colTemp.Item(colTemp.Count + 1 - intLevel)
   End If

End Function
