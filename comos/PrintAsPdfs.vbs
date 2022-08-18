Sub PrintPdf()
  ' attention: bad code, using send keys. not good idea
  
	Set objDocManager = CreateObject("ComosDocumentManager.DocumentManager")
	Set objShell = CreateObject("WScript.Shell")

	strDefaultPrinter = GetDefaultWindowsPrinter()
	'Output strDefaultPrinter
	bSetPrinter = SetDefaultWindowsPrinter("Microsoft Print to PDF")
	'Output bSetPrinter

	If bSetPrinter Then
	
		For i = 1 To b.count
			Set objDoc = b.item(i)
			strFolderPath = "C:\Users\Max\Desktop\test\" & GetCurrentTimestampAsString() & ".pdf"
			objShell.SendKeys strFolderPath
			objShell.SendKeys "{Enter}"
			bExported = objDocManager.PrintDocs(objDoc,1)
		Next
	End If
	
End Function


Function GetDefaultWindowsPrinter()
	GetDefaultWindowsPrinter = ""

	' **** get all printers
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Set colPrinters = objWMIService.ExecQuery("Select * From Win32_Printer")

	If colPrinters.count > 0 Then
		For Each objPrinter In colPrinters
			If objPrinter.Attributes and 4 Then
				GetDefaultWindowsPrinter = objPrinter.Name
			End If
		Next
	End If	
	
End Function


Function SetDefaultWindowsPrinter(strPrinterName)
' Description: Set default windows printer

	SetDefaultWindowsPrinter = false
	
	Set listPrinters = GetWindowsPrinters()
	If listPrinters.Contains(strPrinterName) = false Then Exit Function
	
	Set objNetwork = CreateObject("WScript.Network")
	objNetwork.SetDefaultPrinter strPrinterName
	
	SetDefaultWindowsPrinter = true
	
End Function


Function GetWindowsPrinters()
' Description: returns all windows printer in an arraylist

	Set listPrinters = CreateObject("System.Collections.Arraylist")
	
	' **** get all printers
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Set colPrinters = objWMIService.ExecQuery("Select * From Win32_Printer")

	If colPrinters.count > 0 Then
		For Each objPrinter In colPrinters
			listPrinters.add objPrinter.name
		Next
	End If
	
	Set GetWindowsPrinters = listPrinters
	
End Function


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













'strObject = project.workset.POptions.getoption(7)
'Set objPrinter = CreateObject(strObject)



'Set objDoc = a
'Set objReport = objDoc.Report
'Set objDocument = objReport.ComosDocument

'Output objDocument.GetVersion()


'Set objC = CreateObject("Comos.DocumentTransfer.DocumentTransfer")
'Set colDocs = Project.Workset.GetTempCollection
'colDocs.add a
'Output objC.PrintToPdf(colDocs, "C:\Users\Max\Desktop\super", false, false)

'Set objC = CreateObject("Comos.DocumentTransfer.DocumentExtension")


'Set objPrinter = CreateObject("ComosPrinter.PrinterSets")
'Set objPrinter = CreateObject("Comos.PrintLib.PrinterManager")
'Set objOldPrinter = objPrinter.GetPrinterByName("Microsoft Print to PDF")
'Output objOldPrinter.name
'Output objPrinter.SetDefaultPrinter("Microsoft Print to PDF")


'Set objPrinter = CreateObject("ComosDefRevPrn.AcrobatDistiller")
'Output objPrinter.isprinteravailable
'asdf = objPrinter.DoPrint (a, Nothing, "C:\Users\Max\Desktop\", "asdf.pdf")
'objPrinter.SetPrnDefaults


'Set objRevisionMaster = CreateObject("ComosRevisionMaster.RevisionMaster")
'Set objRevisionMaster.Document = a

'Set objRevisionPrinter = objRevisionMaster.CurrentRevisionPrinter
'objRevisionPrinter.InitPrnSettings

'If objRevisionPrinter.IsPrinterAvailable Then
'	objRevisionPrinter.doprint a, Nothing, "C:\Users\Max\Desktop\", "asdf.pdf"
'Else
'	Output "noting"
'End If


