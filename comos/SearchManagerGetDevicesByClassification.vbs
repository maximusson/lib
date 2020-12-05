' example
Set colDevs = SearchManagerGetDevicesByClassification(a, "A10.A20")
Output colDevs.count


Function SearchManagerGetDevicesByClassification(objStart, strClassificationSearchString)
' DESCRIPTION : uses search manager to get collection of objects under a root node

' SCRIPT REVISIONS :
' 1 - 26-Nov-2020 - created

' INPUT :
' (1) objStart: object from comos tree - [comos object]
' (2) strClassificationString: classification string - [string]

' OUTPUT :
' (1) returns collection of found objects  [boolean]

  Set ws = Project.Workset

  Set SearchManagerGetDevicesByClassification = ws.GetTempCollection
  If objStart Is Nothing Then Exit Function
     
  Set searchManager = ws.GetSearchManager
  Set rootObjects = searchManager.RootObjects
  rootObjects.add objStart
  searchManager.SystemType = 8
  searchManager.AppendSearchCondition "","CLASSIFICATION","4","LIKE", strClassificationSearchString

  Set resultSet = searchManager.Start
  searchManager.RetrieveData(0)
  searchManager.Stop
  Set SearchManagerGetDevicesByClassification = resultSet
       
End Function
