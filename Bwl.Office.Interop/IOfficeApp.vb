Public Interface IOfficeApp
    Property Visible As Boolean
    Sub Close()
    Sub CloseDocument()
    Sub Create()
    Sub CreateDocument()
    Sub FindAndReplace(findText As String, replaceText As String)
    Sub OpenDocument(path As String)
    Sub OpenDocumentCopy(originalPath As String, copyPath As String)
    Sub SaveDocument(Optional newPath As String = "")
End Interface
