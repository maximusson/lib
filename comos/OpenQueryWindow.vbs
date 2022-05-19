'Example:
Set objQuery = a
Output OpenQueryWindow(objQuery)

Function OpenQueryWindow(objQuery)
' DESCRIPTION : opens a COMOS query window
   
' SCRIPT REVISIONS :
' 1 - 18-May-2022 - created
   
' INPUT :
' (1) objQuery: query - [comos object]
   
' OUTPUT :
' (1) returns true if script ran completely [boolean]
   
   OpenQueryWindow = false
   
   If objAttr Is Nothing Then Exit Function
   If objAttr.SystemType <> 2 Then Exit Function
   
   Set ws = Project.Workset
   controlName = ws.Lib.Device.GetClassicTQBProgIdByControlType(ws.Lib.Device.GetControlType(objQuery))
   Call ws.Globals.Application.ShowPropertiesOnMdiChild(objQuery, False, "", "CONTROLTYPE", controlName, Nothing)
         
   OpenQueryWindow = true
         
End Function
