' EXAMPLE :
Set objAttr = a
Output ClickOnButton(objAttr)

Function ClickOnButton(objAttr)
' DESCRIPTION : simulates a click on a button in COMOS gui
   
' SCRIPT REVISIONS :
' 1 : 13-Feb-2020 : created
   
' INPUT :
' (1) objComos: object from comos tree - [comos object]
   
' OUTPUT :
' (1) ClickOnButton: returns true if script ran completely [boolean]
   
        ClickOnButton = false
   
        If objAttr Is Nothing Then Exit Function
        If objAttr.SystemType <> 10 Then Exit Function
        If objAttr.ControlType <> "ComosSUIButton.SUIButton" Then Exit Function
        objAttr.ScriptEngine.ScriptObject.OnClick()
            
		ClickOnButton = true 
            
End Function
