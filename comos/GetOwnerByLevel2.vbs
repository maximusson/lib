'Example Code
Set objComos = GetOwnerByLevel2(a, 3)

If Not objComos Is Nothing Then
   Output objComos.systemfullname
Else
   Output "nothing"
End If


Function GetOwnerByLevel2(ByVal objStart, intLevel)
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

   Set objNavigator = CreateObject("ComosObjNavigator.ObjNavigator")
   objNavigator.AddStep 4, intLevel
   Set GetOwnerByLevel2 = objNavigator.Execute(objStart)

End Function