' EXAMPLE :
Set objQuery = a
strColumnName = "ZT1.ZA1"
bEditable = true
bAdd = SetQueryColumnEditable(objQuery, strColumnName, bolEditable)

Function SetQueryColumnEditable(objQuery, strColumnName, bolEditable)
' DESCRIPTION : sets a query column editable / not editable via script - UNTESTED

' SCRIPT REVISIONS :
' (1) 11-Jul-2022 : Created

' INPUT :
' (1) objQuery : query object - [comos query object]
' (2) strColumnName : column name in query - [name]
' (3) bolEditable : true / false - [boolean]

' OUTPUT :
' (1) SetQueryColumnEditable: returns true if script ran completely [boolean]

	SetQueryColumnEditable = false

	Set objQueryColumn = objQuery.TopQuery.Query.BaseQuery.Columns.Item(strColumnName)
	if objQueryColumn is Nothing Then Exit Function
		
	objQueryColumn.Editable = bolEditable  
  
	SetQueryColumnEditable = true

End Function
