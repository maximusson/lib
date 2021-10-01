Set objQuery = a
Set colColumns = objQuery.XObj.TopQuery.Query.BaseQuery.Columns
strColumnName = "Y00T00156.Y00A00541"
Set objNewColumn = colColumns.addNew(strColumnName, -1)
objNewColumn.description = "Type"

objNewColumn.ScriptTextFunctionObject = "Set ColumnObject = RefColObject.spec(""Y00T00156.Y00A00541"")"
'objNewColumn.ScriptTextFunctionValue = "ColumnValue = ColumnObject.Description"

objNewColumn.ShowProperty = "30"
objNewCOlumn.Editable = True


objQuery.saveAll
