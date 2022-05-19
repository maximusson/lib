Function OnMenuCreate(Popup,Context)
'Event : before creation of context menu
'Input : Context -> Context menu object
'Context -> Context object from which the call is made.
'Context.ComosObject -> Current object or Context.ComosObjects -> Current objects

	' EXAMPLE :
	bAdd = AddPopup(Popup, "New Entry", "ID_NEWENTRY")

End Function

Function AddPopup(Popup, strContextText, strContextID)
' DESCRIPTION : adds an entry to a context menu

' SCRIPT REVISIONS :
' (1) 19-Feb-2019 : created
' (2) 19-May-2022 : beautify script

' INPUT :
' (1) Popup: Popup object - [comos system object]
' (2) strContextText: text that appears within context menu - [string]
' (3) strContextID: id for context menu entry, make sure it is unique - [unique string]

' OUTPUT :
' (1) AddPopup: returns true if script ran completely [boolean]

 	AddPopup = false
 	Popup.add strContextText, strContextID
 	AddPopup = true

End Function
