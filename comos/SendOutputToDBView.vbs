' EXAMPLE :
strMessage = "Super Message."
bSent = SendOutputToDBView(strMessage)

Function SendOutputToDBView(strMessage)
' DESCRIPTION : send message to debug outpt, that can be read from dbview.exe

' SCRIPT REVISIONS :
' (1) 20-May-2022 : created

' INPUT :
' (1) strMessage: text that contains message for output with dbview - [string]]

' OUTPUT :
' (1) SendOutputToDBView: returns true if script ran completely [boolean]

	SendOutputToDBView = false
	
  Project.Workset.Lib.Output strMessage
	
	SendOutputToDBView = true

End Function
