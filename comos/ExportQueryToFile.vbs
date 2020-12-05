'Example:
Set objQuery = a
strFilepath = "C:\temp\ExportData.xlsx"
strExcelSheetName = "test"
Output ExportQueryToFile(objQuery, strFilepath, strExcelSheetName)

Function ExportQueryToFile(objQuery, strFilepath, strExcelSheetName)
' DESCRIPTION : exports query to file on filesystem.
' valid file types: txt, xls, xlsx, xml, mdb

' SCRIPT REVISIONS :
' 1 - 10-Mar-2019 - Created

' INPUT
' (1) objQuery: query object - [comos object]
' (2) strFilepath: path on filesystem for export - [string]
' (3) strExcelSheetName: name of sheet, when exported to excel (use empty string if not used) - [string]

' OUTPUT :
' (1) returns true if script ran completely [boolean]
    ExportQueryToFile = false

    If IsFilepathUniqueAndFolderpathValid(strFilepath) = false Then Exit Function

    strExtension = GetFileExtension(strFilepath)
    If strExtension = "" Then Exit Function

    Select Case strExtension
    Case "txt"
      intExportDataType = 0
    Case "xls"
      intExportDataType = 1
    Case "xlsx"
      intExportDataType = 1
    Case "xml"
      intExportDataType = 2
    Case "mdb"
      intExportDataType = 3
    Case Else
      Exit Function
    End Select

    Set objTQ = objQuery.XObj.TopQuery
    objTQ.Execute
    objTQ.Query.ExportData intExportDataType, strFilepath, strExcelSheetName

    ExportQueryToFile = true

End Function


Function IsFilepathUniqueAndFolderpathValid(strFilepath)
' DESCRIPTION: checks if filepath is unique. scripts is cancelled if file already exists.
' checks if folder exists. scripts is cancelled if folder does not exist

    IsFilepathUniqueAndFolderpathValid = false

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Check if file exists, exit function if true
    If fso.FileExists(strFilepath) Then Exit Function

    ' Check if folder exists, exit function if false
    arrStr = split(strFilepath, "\")
    If UBound(arrStr) = 0 Then Exit Function
    strFolderpath = left(strFilepath,len(strFilepath)-len(arrStr(UBound(arrStr))))
    If fso.FolderExists(strFolderpath) = false Then Exit Function

    IsFilepathUniqueAndFolderpathValid = true

End Function


Function GetFileExtension(strFilepath)
' DESCRIPTION: returns file extension from a given filepath

    GetFileExtension = ""

    Set fso = CreateObject("Scripting.FileSystemObject")

    arrStr = split(strFilepath, ".")

    ' Check if at least one dot appears in path, exit function if not
    If UBound(arrStr) = 0 Then Exit Function

    GetFileExtension = LCase(arrStr(UBound(arrStr)))

End Function
