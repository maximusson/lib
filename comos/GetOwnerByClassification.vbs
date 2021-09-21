'Example:
intSearchKey = 4
strSearchClassification = "PM.C10"
Set objStart = a
bolCheckStartObjectItself = true
Set objOwner = GetOwnerByClassification(intSearchKey, strSearchClassification, objStart, bolCheckStartObjectItself)


Function GetOwnerByClassification(intSearchKey, strSearchClassification, ByVal objStart, bolCheckStartObjectItself)
' DESCRIPTION : searches owner by classfication and returns it if found - otherwise nothing

' SCRIPT REVISIONS :
' 1 - 01-Aug-2019 - Created

' INPUT :
' (1) intSearchKey: classification number as integer, 1, 2, 3 or 4 - [integer]
' (2) strSearchClassification: classification search string - "PM.B10" - [string]
' (3) objStart: root object, [object]
' (4) bolCheckStartObjectItself: boolean value, that checks if root object itselt should be checked - [boolean]

' OUTPUT :
' (1) owner with specified classification, if no owner exists then nothing - [object]
   Set GetOwnerByClassification = Nothing

   If objStart Is Nothing Then Exit Function

   ' Check if object itself is classified with key
   If bolCheckStartObjectItself = true Then
      If objStart.ClassificationExists(intSearchKey,strSearchClassification) Then
         Set GetOwnerByClassification = objStart
         Exit Function
      End If
   End If

   ' Loop over all owner and check for classification
   Do While Not objStart.owner Is Nothing
      Set objStart = objStart.owner
      If objStart.SystemType <> intSystemtype Then Exit Function
      If objStart.ClassificationExists(intSearchKey,strSearchClassification) Then
         Set GetOwnerByClassification = objStart
         Exit Function
      End If
   Loop

End Function
