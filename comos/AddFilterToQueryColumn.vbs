' EXAMPLE :
Set objQuery = a
strColumnName = "Column1"
strFilterValue = "100"
bFilterAdded = AddFilterToQueryColumn(objQuery, strColumnName, strFilterValue)

Function AddFilterToQueryColumn(objQuery, strColumnName, strFilterValue)
' DESCRIPTION : adds a filter to a query column
   
' SCRIPT REVISIONS :
' (1) 18-May-2022 : created
' (2) 19-May-2022 : beautify script
   
' INPUT :
' (1) objQuery: query - [comos query object]
' (2) strColumnName: name of column in query - [string]
' (3) strFilterName: value for filter - [string]
   
' OUTPUT :
' (1) AddFilterToQueryColumn: returns true if script ran completely [boolean]
   
	OpenQueryWindow = false
   
	If objQuery Is Nothing Then Exit Function
	If objQuery.SystemType <> 2 Then Exit Function
         
	Set ws = Project.Workset
	Set objNewFilter = objQuery.Filter.AddNew
	Set objNewFilter.Column = objQuery.BaseQuery.Columns.Item(strColumnName)
	objNewFilter.Value = strFilterValue
	objNewFilter.Operator = 9
         
	OpenQueryWindow = true 
         
End Function

		
' Set Query = a
' Set objTopQuery = Query.XObj.TopQuery 
' objTopQuery.Execute 
' 
' set filter
' Set objQueryFilter = objTopQuery.Query.Filter
' objQueryFilter.Clear
' objQueryFilter.AddNew
' Set objQueryFilter.Item(1).Column = Query.BaseQuery.Columns.Item(strColumnName)
' objQueryFilter.Item(1).Value = strSearchString
' objQueryFilter.Item(1).Operator = 10 ' qcVBLike
' objQueryFilter.Item(1).CaseSensitive = False 
' Query.Refresh
' QueryBrowser.Refresh
' 		
' Set objQueryFilter = Nothing
' Set objTopQuery = Nothing
