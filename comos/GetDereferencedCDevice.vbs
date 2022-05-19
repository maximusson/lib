' EXAMPLE :
Set objComos = a
Set objDerefCDevice = GetDereferencedCDevice(objComos)

Function GetDereferencedCDevice(objComos)
' DESCRIPTION: returns the very end of cdevices

' VERSIONS:
' (1) 18-May-2022 : created

' INPUTS:
' (1) objComos: comos object of systemtype 8 or 13 [comos object]

' OUTPUS:
' (1) GetDereferencedCDevice: returns the very end of cdevice collection

	Set GetDereferencedCDevice = Nothing
	If objComos Is Nothing Then Exit Function
	Set objCDevice = Nothing

	Select Case objComos.SystemType
	Case 8
		' Device
		Set objCDevice = objComos.CDevice
	Case 13
		' CDevice
		Set objCDevice = objComos
	Case Else
		Exit Function
	End Select

	If objCDevice Is Nothing Then Exit Function
   
	Set GetDereferencedCDevice = objCDevice
	Do While Not GetDereferencedCDevice.CDevice Is Nothing
		Set GetDereferencedCDevice = GetDereferencedCDevice.CDevice
	Loop

End Function
