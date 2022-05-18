'Example Code
Set objComos = a
Set objDerefCDevice = GetDereferencedCDevice(objComos)


Function GetDereferencedCDevice(objComos)
' DESCRIPTION: returns the very end of cdevices

' VERSIONS:
' 1 - 18-May-2022 - created

' INPUTS:
  ' (1) objComos - comos object of systemtype 8 or 13 [comos object]

' OUTPUS:
' (1) returns the very end of cdevice collection

   Set GetDereferencedCDevice = Nothing
   If objComos Is Nothing Then Exit Function
   Set objCDev = Nothing

   Select Case objComos.SystemType
      Case 8
         ' Device
         Set objCDev = objComos.CDevice
      Case 13
         ' CDevice
         Set objCDev = objComos
      Case Else
	 exit Function
   End Select

   If objCDev Is Nothing Then Exit Function
   
   Set GetDereferencedCDevice = objCDev
   Do While Not GetDereferencedCDevice.CDevice Is Nothing
       Set GetDereferencedCDevice = GetDereferencedCDevice.CDevice
   Loop

End Function
