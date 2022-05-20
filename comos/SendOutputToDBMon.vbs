' EXAMPLE :
strMessage = "Super Message."
bSent = SendOutputToDBMon(strMessage)

Function SendOutputToDBMon(strMessage)
' DESCRIPTION : send message to dbmon, that can be read from dbmon.exe

' SCRIPT REVISIONS :
' (1) 20-May-2022 : created

' INPUT :
' (1) strMessage: text that contains message for output with dbmon - [string]]

' OUTPUT :
' (1) SendOutputToDBMon: returns true if script ran completely [boolean]

	SendOutputToDBMon = false
	
	Set XStdMod = CreateObject("ComosXStdMod.XStdMod")
	XStdMod.Outputdebugstring strMessage
	
	SendOutputToDBMon = true

End Function
