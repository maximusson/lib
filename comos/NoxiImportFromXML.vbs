' EXAMPLE :
Set objRoot = a
strFilepath = "C:\NoxiExport.xml"
output NoxiImportFromXML(objRoot, strFilepath)

Function NoxiImportFromXML(objRoot, strFilepath)
' DESCRIPTION : imports comos object(s) from xml file

' SCRIPT REVISIONS :
' 1 : 01-May-2022 : Created
' 2 : 19-May-2022 : beautify script

' INPUT :
' (1) objRoot: object from comos tree - [comos object]
' (2) strFilepath: path on filesystem - [string]

' OUTPUT :
' (1) NoxiImportFromXML: returns true if script ran completely [string]
  
	NoxiImportFromXML = false
  
	If objRoot Is Nothing Then Exit Function

	Set objNoxi = CreateObject("ComosNOXIE.Noxie")
	objNoxi.DeSerializeToFile objRoot, strFilepath
	Set objNoxi = Nothing

	NoxiImportFromXML = true
    
End Function
