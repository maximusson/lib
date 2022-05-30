' ############### GROUP ###########################
Output "GROUP ASSIGNMENTS"
Set objUser = Project.Workset.LoadObjectByType(33, "A478X5FTXW")

Set colCLinks = objUser.GetBackPointerCLinksWithReference
For i = 1 To colCLinks.count
	Set objCLink = colCLinks.Item(i)
	Set objGroup = objCLink.Owner
	
	Output "group: " & objGroup.Name
	Output "assigned user: " & objUser.Name 
	Output "assigned on: " & objCLink.CS.Login
	Output "assigned by: " & objCLink.CS.User.Name
	Output " "
Next


' ################ RIGHTS #########################
Output "RIGHT ASSIGNMENTS"
Set objUser = Project.Workset.LoadObjectbyType(33, "A2ST1EP6IA")

Set colRights = objUser.AllBackpointerRights
For i = 1 To colRights.count
	Set objRight = colRights.item(i)
	
	Output "reference system type name: " & objRight.Reference.SystemTypeName
	Output "reference name and description: " & objRight.Reference.name & " " & objRight.Reference.Description
	Output "assigned user/group: " & objUser.name
	Output "assigned on: " & objRight.CS.Login
	Output "assigned by: " & objRight.CS.User.Name	
	Output ""
Next
