'Example:
Set objComos = a
Output OpenPropertiesWindowOfObject(a)

Function OpenPropertiesWindowOfObject(objComos)
' DESCRIPTION : opens properties window of a COMOS object

' SCRIPT REVISIONS :
' 1 - 28-Feb-2020 - Created

' INPUT :
' (1) objComos: object from comos tree - [comos object]
' (2) strFilepath: path on filesystem - [string]

' OUTPUT :
' (1) returns true if script ran completely [boolean]
    OpenPropertiesWindowOfObject = false

    If objComos Is Nothing Then Exit Function

    Set objNavi = Project.Workset.Globals.NAVIGATOR
    objNavi.GetCurrentTree.DefaultAction objComos

    OpenPropertiesWindowOfObject = true

End Function
