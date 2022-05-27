' EXAMPLE :
Set objCTab = a
Set objCAttribute = b
intPositionX = 2
intPositionY = 2
bPlaced = PlaceAttributeOnTab(objCTab, objCAttribute, intPositionX, intPositionY)
Output bPlaced

Function PlaceAttributeOnTab(objCTab, objCAttribute, intPositionX, intPositionY)
' DESCRIPTION : places an attribute on a tab

' SCRIPT REVISIONS :
' (1) 27-May-2022 : Created and successfully tested

' INPUT :
' (1) objCTab: comos object of new owner - [comos cspecification tab]
' (2) objCAttribute: comos object - template for copy - [comos cspecification attribute]

' OUTPUT :
' (1) PlaceAttributeOnTab: returns true if placing was successful [boolean]

	PlaceAttributeOnTab = false
	
	If objCTab.SystemType <> 10 or objCAttribute.Systemtype <> 10 Then Exit Function
	If intPositionX < 0 or intPositionY < 0 Then Exit Function
	
	Set ws = Project.Workset
	
	If objCTab.OwnSpecifications.ItemExist(objCAttribute.Name) = True Then Exit Function
	
	Set objNewAttribute = objCTab.OwnSpecifications.CreateNew
	objNewAttribute.CSpecification = objCAttribute
	objNewAttribute.Name = objCAttribute.Name
	ws.lib.sui.CtrlProperty (27,objNewAttribute) = intPositionX
	ws.lib.sui.CtrlProperty (28,objNewAttribute) = intPositionY
	objNewAttribute.Save
	
	' Return
	PlaceAttributeOnTab = true
	
End Function
