Sub SaveOutlookAttachmentsOnFileSystem()

    strPath = "C:\temp\"
    
    For Each objMail In Application.ActiveExplorer.Selection
        If TypeName(objMail) = "MailItem" Then
            For Each objAttachment In objMail.Attachments
                Debug.Print objAttachment.FileName
                objAttachment.SaveAsFile strPath & objAttachment.FileName
            Next
        End If
    Next

End Sub