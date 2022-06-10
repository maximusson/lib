' EXAMPLE :
bRefreshed = RefreshTab()

Function RefreshTab(objTab)
' DESCRIPTION : refresh tab - UNTESTED

' SCRIPT REVISIONS :
' (1) 10-Jun-2022 : Created

' INPUT :
' (1) objTab: tab object - [comos tab object]

' OUTPUT :
' (1) RefreshTab: returns true if script ran completely [boolean]

	RefreshTab = false
  Set ws = Project.Workset
  
  If objTab.SystemType <> 10 Then Exit Function
    
  ws.Globals.AppCommand.RefreshDevice
  ws.Lib.RefreshCurrentChapterBySpecOwner(objTab)
    
	RefreshTab = true

End Function
