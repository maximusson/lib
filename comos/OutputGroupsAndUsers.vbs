' example:
Call OutputGroupsAndUsers()

Sub OutputGroupsAndUsers()
' DESCRIPTION: outputs all groups and group users
  Set ws = Project.Workset
  Set colComosGroups = ws.GetAllUserGroups
  output "System Uid" & vbTab & "Name" & vbTab & "Description"
  For i = 1 To colComosGroups.count
    Set objGroup = colComosGroups.item(i)
    If Not objGroup Is Nothing Then
      Set colMembers = objGroup.AllUsers
      For j = 1 To colMembers.count
        Set objMember = colMembers.item(j)
        If Not objMember Is Nothing Then
          strGroup = objGroup.SystemUid & vbTab & objGroup.Name & vbTab & objGroup.Description
          strMember = objMember.SystemUid & vbTab & objMember.Name & vbTab & objMember.Description
          Output strGroup & vbTab & strMember              
        End If
      Next     
    End If
  Next
End Sub
