' EXAMPLE :
Set objComos = a
strFilepath = "C:\NoxiExport.xml"
output NoxiExportToXML(objComos, strFilepath)

Function NoxiExportToXML(objComos, strFilepath)
' DESCRIPTION : exports comos object(s) to xml file

' SCRIPT REVISIONS :
' (1) 01-Aug-2019 : created
' (2) 19-May-2022 : beautify script
	
' INPUT :
' (1) objComos: object from comos tree - [comos object]
' (2) strFilepath: path on filesystem - [string]

' OUTPUT :
' (1) NoxiExportToXML: returns true if script ran completely [string]
   
	NoxiExportToXML = false

	If objComos Is Nothing Then Exit Function

	' Get Noxi Manager from dll
	Set noxi = CreateObject("ComosNOXIE.Noxie")
	strSerialize = noxi.SerializeToFile(objComos, strFilepath)
	Set noxi = Nothing

	NoxiExportToXML = true
      
End Function
