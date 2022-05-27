' EXAMPLE :
Set objQuery = a
bOpened = OpenQueryWindow(objQuery)

Function OpenQueryWindow(objQuery)
' DESCRIPTION : opens a COMOS query window
   
' SCRIPT REVISIONS :
' (1) 18-May-2022 : created
' (2) 27-May-2022 : removed check for systemtype, successfully tested
	
' INPUT :
' (1) objQuery: query - [comos object]
   
' OUTPUT :
' (1) returns true if script ran completely [boolean]
   
	OpenQueryWindow = false
   
	If objQuery Is Nothing Then Exit Function
   
	Set ws = Project.Workset
	controlName = ws.Lib.Device.GetClassicTQBProgIdByControlType(ws.Lib.Device.GetControlType(objQuery))
	Call ws.Globals.Application.ShowPropertiesOnMdiChild(objQuery, False, "", "CONTROLTYPE", controlName, Nothing)
         
	OpenQueryWindow = true
         
End Function
