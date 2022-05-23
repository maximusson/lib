' EXAMPLE :
strLogFilePath = "C:\logfile.txt" 
bLogged = WriteLogToFile(strLogFilePath, 0, "", "", Nothing) ' create text file
bLogged = WriteLogToFile(strLogFilePath, 1, "", "", Nothing) ' write header
bLogged = WriteLogToFile(strLogFilePath, 2, "information", "data ok", Nothing) ' write information
bLogged = WriteLogToFile(strLogFilePath, 2, "warning", "attribute not changed", a) ' write information
bLogged = WriteLogToFile(strLogFilePath, 2, "error", "object could not be found", b) ' write information
    
Function WriteLogToFile(strLogFilePath, intOption, strActionTitle, strMessage, objComos)
' DESCRIPTION : writes information to a logfile - UNTESTED

' SCRIPT REVISIONS :
' (1) 20-May-2022 : created

' INPUT :
' (1) strLogFilePath: path for new logfile - [string]
' (2) intOption: integer - 0, 1 or 2 - choose 0 for creating logfile, 1 for writing to header, 2 for writing information [string]
' (3) intActionTitle: choose a title for your log (only effective with intOption = 2) [string]
' (4) intMessage: choose a message for your log (only effective with intOption = 2) [string]
' (5) objComos: log current comos object to your file [comos object]

' OUTPUT :
' (1) WriteLogToFile: returns true if script ran completely [boolean]

	WriteLogToFile = False
    
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set ws = Project.Workset
	strHeader = "date" & vbTab & "time" & vbTab & "action" & vbTab & "message" & vbTab & "object systemfullname" & vbTab & "user" & vbTab & "project" & vbTab & "overlay"
        
	Select Case intOption
	Case 0 ' create / override logfile
		If fso.FileExists(strLogFilePath) Then
			Set objFile = fso.OpenTextFile(strLogFilePath, 2)
		Else
			Set objFile = fso.CreateTextFile(strLogFilePath)
		End If
		objFile.Close
    
	Case 1 ' write header
		If fso.FileExists(strLogFilePath) = False Then Exit Function
		Set objFile = fso.OpenTextFile(strLogFilePath, 8)
		objFile.WriteLine strHeader
		objFile.Close
        
	Case 2 ' write information
		If fso.FileExists(strLogFilePath) = False Then Exit Function
		Set objCurrentUser = ws.GetCurrentUser
		Set objCurrentProject = ws.GetCurrentProject
		Set objCurrentOverlay = objCurrentProject.CurrentWorkingOverlay
		strUser = objCurrentUser.Description
		strProject = objCurrentProject.Name & " " & objCurrentProject.Description
		strOverlay = ""
		If Not objCurrentOverlay Is Nothing Then
			strOverlay = objCurrentOverlay.Name & " " & objCurrentOverlay.Description
		End If
		strSystemfullname = ""
		If Not objComos Is Nothing Then
			strSystemfullname = objComos.Systemfullname
		End If
        
		strData = Replace(strData, "date", Year(Now) & "-" & Right("00" & Month(Now), 2) & "-" & Right("00" & Day(Now), 2)) & vbTab & _
		Replace(strData, "time", Right("00" & Hour(Now), 2) & ":" & Right("00" & Minute(Now), 2) & ":" & Right("00" & Second(Now), 2)) & vbTab & _
		Replace(strData, "action", strActionTitle) & vbTab & _
		Replace(strData, "message", strMessage) & vbTab & _
		Replace(strData, "object systemfullname", "strSystemfullname") & vbTab & _
		Replace(strData, "user", "strUser") & vbTab & _
		Replace(strData, "project", "strProject") & vbTab & _
		Replace(strData, "overlay", "strOverlay")
        
		Set objFile = fso.OpenTextFile(strLogFilePath, 8)
		objFile.WriteLine strDate
		objFile.Close
        
	Case Else
		Exit Function
				
	End Select

	WriteLogToFile = True

End Function
