' EXAMPLE :
Set objAttribute = a
bClicked = ClickOnButton(objAttribute)

Function ClickOnButton(objAttribute)
' DESCRIPTION : simulates a click on a button in COMOS gui
   
' SCRIPT REVISIONS :
' 1 : 13-Feb-2020 : created
' 2 : 19-May-2022 : beautify
   
' INPUT :
	' (1) objComos: object from comos tree - [comos attribute object]
   
' OUTPUT :
' (1) ClickOnButton: returns true if script ran completely [boolean]
   
	ClickOnButton = false
   
	If objAttribute Is Nothing Then Exit Function
	If objAttribute.SystemType <> 10 Then Exit Function
	If objAttribute.ControlType <> "ComosSUIButton.SUIButton" Then Exit Function
				
	objAttribute.ScriptEngine.ScriptObject.OnClick()
            
	ClickOnButton = true 
            
End Function
