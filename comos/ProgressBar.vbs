' EXAMPLE :
Call ProgressBar()

Sub ProgressBar()
' DESCRIPTION: example script for usage of progressbar - UNTESTED
  
' REVISION:
' (1) 03-Nov-2021 : created
' (2) 19-May-2022 : beautify script
' (3) 15-Jun-2022 : typo corrected
	
' INPUTS:
' () 

' OUTPUS:
' () 	
	
	' start progressbar
	Set objProgressBar = CreateObject("ComosXMLContent.Progress")
  	objProgressBar.Caption = "Running ... "
	objProgressBar.Percentage = 1
	objProgressBar.Where = 1
	objProgressBar.StartProcess
	
  	' loop
  	intCount = 10000
  	For i = 1 To intCount
		
		' do stuff
    
		' update progressbar
		If objProgressbar.State = 3 Then Exit Sub
		objProgressBar.Percentage = Round(i / intCount * 100)	
		
	Next
	
	' end progressbar
	objProgressBar.StopProcess
    
End Sub
