'Example:
Set objQuery = a
strColumnName = "Column1"
strFilterValue = "100"
Output AddFilterToQueryColumn(objQuery, strColumnName, strFilterValue)

Function AddFilterToQueryColumn(objQuery, strColumnName, strFilterValue)
' DESCRIPTION : adds a filter to a query column
' SCRIPT REVISIONS :
' 1 - 18-May-2022 - created
' INPUT :
' (1) objQuery: query - [comos object]
' (2) strColumnName: name of column in query - [string]
' (3) strFilterName: value for filter - [string]
' OUTPUT :
' (1) returns true if script ran completely [boolean]
   OpenQueryWindow = false
   If objAttr Is Nothing Then Exit Function
   If objAttr.SystemType <> 2 Then Exit Function
   Set ws = Project.Workset
   Set objNewFilter = objQuery.Filter.AddNew
   Set objNewFilter.Column = Query.BaseQuery.Columns.Item(strColumnName)
   objNewFilter.Value = strFilterValue
   objNewFilter.Operator = 9
   OpenQueryWindow = true
End Function
