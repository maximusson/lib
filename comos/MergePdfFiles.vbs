' EXAMPLE :
Set listFilepaths = CreateObject("System.Collections.Arraylist")
listFilepaths.add "C:\test1.pdf"
listFilepaths.add "C:\test2.pdf"
strFilepath = "C:\merged.pdf"
bMerged = MergePdfFiles(listFilepaths, strFilepath)

Function MergePdfFiles(listFilepaths, strFilepath)
' DESCRIPTION: merges multiple pdfs into one pdf - ToDo: check if export path is valid, check if pdf file exists

' VERSIONS:
' (1) 18-May-2022 : created
' (2) 19-May-2022 : beautify script
' (3) 28-Feb-2023 : corrected quickPdf -> to objQuickPdf
	
' INPUTS:
' (1) listFilepaths: arraylist of filepaths [arraylist] - CreateObject("System.Collections.Arraylist")
' (2) strFilepath: filepath for merged pdf [string]

' OUTPUS:
' (1) MergePdfFiles: status whether export was succesful or not [boolean]

	MergePdfFiles = false

	Set objQuickPdf = CreateObject("QuickPDFAX0812.PDFLibrary")
	objQuickPDF.UnlockKey("jt9593uh8eh5ai4cu9b36hb5y")
	strFileListName = "FilesToMerge"
	objQuickPdf.ClearFileList fileListName
	
	For each strPdfFilepath in listFilepaths
		objQuickPdf.AddToFileList strFileListName, strPdfFilepath	
	Next

	objQuickPdf.MergeFileListFast strFileListName, strFilepath

	MergePdfFiles = true

End Function
