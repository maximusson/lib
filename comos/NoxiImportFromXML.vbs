'Example:
Set objRoot = a
strFilepath = "C:\NoxiExport.xml"
output NoxiImportFromXML(objRoot, strFilepath)


Function NoxiImportFromXML(objRoot, strFilepath)
' DESCRIPTION : imports comos object(s) from xml file

' SCRIPT REVISIONS :
' 1 - 01-May-2022 - Created

' INPUT :
' (1) objRoot: object from comos tree - [comos object]
' (2) strFilepath: path on filesystem - [string]

' OUTPUT :
  ' (1) returns true if script ran completely [string]
  NoxiImportFromXML = false
  
  If objRoot Is Nothing Then Exit Function

  ' Get Noxi Manager from dll
  Set noxi = CreateObject("ComosNOXIE.Noxie")
  noxi.DeSerializeToFile objRoot, strFilepath
  Set noxi = Nothing

  NoxiImportFromXML = true
    
End Function
