' EXAMPLE :
Set objComos = a
Output NavigateToObjectInComosTree(a)

Function NavigateToObjectInComosTree(objComos)
' DESCRIPTION : navigate to any comos object within a tree

' SCRIPT REVISIONS :
' 1 - 10-Jan-2020 - Created

' INPUT :
' (1) objComos: object from comos tree - [comos object]

' OUTPUT :
' (1) NavigateToObjectInComosTree: returns true if script ran completely [boolean]
	
	NavigateToObjectInComosTree = false

 	If objComos Is Nothing Then Exit Function

 	Set objNavi = Project.Workset.Globals.NAVIGATOR
 	objNavi.pltObject = objComos

 	NavigateToObjectInComosTree = true

End Function
