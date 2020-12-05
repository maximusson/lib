'Example:
Set objComos = a
strFilepath = "C:\NoxiExport.xml"
output NoxiExportToXML(objComos, strFilepath)


Function NoxiExportToXML(objComos, strFilepath)
' DESCRIPTION : exports comos object(s) to xml file

' SCRIPT REVISIONS :
' 1 - 01-Aug-2019 - Created

' INPUT :
' (1) objComos: object from comos tree - [comos object]
' (2) strFilepath: path on filesystem - [string]

' OUTPUT :
' (1) ?? [string]
   NoxiExportToXML = ""

   If objComos Is Nothing Then Exit Function

   ' Get Noxi Manager from dll
   Set noxi = CreateObject("ComosNOXIE.Noxie")
   NoxiExportToXML = noxi.SerializeToFile(objComos, strFilepath)
   Set noxi = Nothing

End Function
