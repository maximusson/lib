' example:
Call Progressbar()

Sub Progressbar()
  ' DESCRIPTION: example script for usage of progressbar
  
  ' REVISION:
  ' 1   03-Nov-2021   created
  
	' start progressbar
	Set objProgressbar = CreateObject("ComosXMLContent.Progress")
  objProgressbar.Caption = "Running ... "
	objProgressbar.Percentage = 1
	objProgressbar.Where = 1
	objProgressbar.StartProcess
	
  ' loop
  intCount = 10000
  For i = 1 To intCount
		
    ' do stuff
    
		' update progressbar
		If objProgressbar.State = 3 Then Exit Sub
    objProgressbar.Percentage = Round(i / intCount * 100)	
		
	Next
	
	' end progressbar
	objProgressbar.StopProcess
    
End Sub
