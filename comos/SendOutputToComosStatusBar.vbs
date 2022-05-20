' EXAMPLE :
strMessage = "Super Message."
bSent = SendOutputToComosStatusBar(strMessage)

Function SendOutputToComosStatusBar(strMessage)
' DESCRIPTION : send message to COMOS status bar

' SCRIPT REVISIONS :
' (1) 20-May-2022 : created

' INPUT :
	' (1) strMessage: text that contains message for status bar - [string]]

' OUTPUT :
' (1) SendOutputToComosStatusBar: returns true if script ran completely [boolean]

	SendOutputToComosStatusBar = false
	
	Set objXStdMod = CreateObject("ComosXStdMod.XStdMod")
	objXStdMod.UserPrompt(strMessage)
	
	SendOutputToComosStatusBar = true

End Function
