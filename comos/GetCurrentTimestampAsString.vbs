' EXAMPLE :
strCurrentTimestamp = GetCurrentTimestampAsString()

Function GetCurrentTimestampAsString()
' DESCRIPTION: returns string from current timestamp, that looks like that: 20210919-151703
	
' SCRIPT REVISIONS :
' (1) 02-Feb-2022 : Created
' (2) 19-May-2022 : beautify script
	
' INPUT :
' () 

' OUTPUT :
' (1) GetCurrentTimestampAsString: current timestamp[string] 
	
	strYear = year(now)
	strMonth = right("00" & month(now),2)
	strDay = right("00" & day(now),2)
	strHour = right("00" & hour(now),2)
	strMinute = right("00" & minute(now),2)
	strSecond = right("00" & second(now),2)
	GetCurrentTimestampAsString = strYear & strMonth & strDay & "-" & strHour & strMinute & strSecond
	
End Function
