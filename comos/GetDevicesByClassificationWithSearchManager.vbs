' EXAMPLE :
Set colDevs = GetDevicesByClassificationWithSearchManager(a, "4", "A10.A20")
Output colDevs.count

Function GetDevicesByClassificationWithSearchManager(objStart, strClassificationKey, strClassificationSearchString)
' DESCRIPTION : uses search manager to get collection of objects under a root node

' SCRIPT REVISIONS :
' (1) 26-Nov-2020 : created
' (2) 19-May-2022: beautify script
' (3) 20-May-2022: renamed function
' (4) 30-May-2022: function restored (accidentally overriden)
	
' INPUT :
' (1) objStart: object from comos tree - [comos object]
' (2) strClassificationKey: key for classification (1, 2, 3 or 4) [string]
' (3) strClassificationSearchString: classification string - [string]

' OUTPUT :
' (1) GetDevicesByClassificationWithSearchManager: returns collection of found objects  [collection]

	Set ws = Project.Workset

	Set SearchManagerGetDevicesByClassification = ws.GetTempCollection
	If objStart Is Nothing Then Exit Function
     
	Set searchManager = ws.GetSearchManager
	Set rootObjects = searchManager.RootObjects
	rootObjects.add objStart
	searchManager.SystemType = 8
	searchManager.AppendSearchCondition "","CLASSIFICATION",strClassificationKey,"LIKE", strClassificationSearchString

	Set resultSet = searchManager.Start
	searchManager.RetrieveData(0)
	searchManager.Stop
	Set GetDevicesByClassificationWithSearchManager = resultSet
       
End Function
