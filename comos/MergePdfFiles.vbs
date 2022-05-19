' EXAMPLE :
Set listFilepaths = CreateObject("System.Collections.Arraylist")
listFilepaths.add "C:\test1.pdf"
listFilepaths.add "C:\test2.pdf"
strFilepath = "C:\merged.pdf"
bMerged = MergePdfFiles(listFilepaths, strFilepath)

Function MergePdfFiles(listFilepaths, strFilepath)
' DESCRIPTION: merges multiple pdfs into one pdf

' VERSIONS:
' 1 : 18-May-2022 : created

' INPUTS:
' (1) listFilepaths: arraylist of filepaths [arraylist] - CreateObject("System.Collections.Arraylist")
' (2) strFilepath: filepath for merged pdf [string]

' OUTPUS:
' (1) MergePdfFiles: status whether export was succesful or not [boolean]

	MergePdfFiles = false

	Set quickPdf = CreateObject("QuickPDFAX0812.PDFLibrary")
	quickPDF.UnlockKey("jt9593uh8eh5ai4cu9b36hb5y")
	strFileListName = "FilesToMerge"
	objQuickPdf.ClearFileList fileListName
	
	For each strPdfFilepath in listFilepaths
		objQuickPdf.AddToFileList strFileListName, strPdfFilepath	
	Next

	objQuickPdf.MergeFileListFast strFileListName, strFilepath

	MergePdfFiles = true

End Function
