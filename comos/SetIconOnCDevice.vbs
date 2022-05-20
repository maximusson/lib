' EXAMPLE :
strIconRelativePath = "icons\@00\@10.ICO"
Set objCDevice = a
bSet = SetIconOnCDevice(objCDevice, strIconRelativePath)

Function SetIconOnCDevice(objCDevice, strIconRelativePath)
' DESCRIPTION : set icon on cdevice - UNTESTED

' SCRIPT REVISIONS :
' (1) 20-May-2022 : created

' INPUT :
' (1) objCDevice: cdevice where icon should be changed - [comos cdevice object]
' (2) strContstrIconRelativePathextText: realtive path of icon - [string]

' OUTPUT :
' (1) SetIconOnCDevice: returns true if script ran completely [boolean]

	SetIconOnCDevice = false
	
	If objCDevice.SystemType <> 13 Then Exit Function
	If strIconRelativePath = "" Then Exit Function
	
	objCDevice.OwnIconFileName = strIconRelativePath
  
	SetIconOnCDevice = true

End Function
