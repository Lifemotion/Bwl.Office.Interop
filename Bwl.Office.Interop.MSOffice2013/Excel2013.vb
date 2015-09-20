Imports Microsoft.Office.Interop.Excel

Public Class Excel2013
    Implements IExcel

    Private _app As Application
    Private _doc As Workbook

    Public Property Visible As Boolean Implements IExcel.Visible
        Get
            CheckCreated()
            Return _app.Visible
        End Get
        Set(value As Boolean)
            CheckCreated()
            _app.Visible = value
        End Set
    End Property

    Public Shared Sub TestWorking()
        Dim app As New Application
        app.Visible = True
        app.Visible = False
        app.Quit()
    End Sub

    Private Sub CheckCreated()
        If _app Is Nothing Then Throw New Exception("Excel Application not created. Use Create method before.")
    End Sub

    Private Sub CheckDocument()
        CheckCreated()
        If _doc Is Nothing Then Throw New Exception("Excel Document not opened. Use OpenDocument method before.")
    End Sub

    Public Sub Create() Implements IExcel.Create
        If _app IsNot Nothing Then Throw New Exception("Excel Application already created.")
        _app = New Application
    End Sub

    Public Sub Close() Implements IExcel.Close
        _app.Quit()
        _app = Nothing
    End Sub

    Public Sub OpenDocument(path As String) Implements IExcel.OpenDocument
        CheckCreated()
        Dim fi As New IO.FileInfo(path)
        If _doc IsNot Nothing Then Throw New Exception("Excel Document is already opened. Ise CloseDocument before open new one.")
        _doc = _app.Workbooks.Open(fi.FullName)
    End Sub

    Public Sub OpenDocumentCopy(originalPath As String, copyPath As String) Implements IExcel.OpenDocumentCopy
        CheckCreated()
        Dim fi As New IO.FileInfo(originalPath)
        fi.CopyTo(copyPath)
        fi = New IO.FileInfo(copyPath)
        If _doc IsNot Nothing Then Throw New Exception("Excel Document is already opened. Ise CloseDocument before open new one.")
        _doc = _app.Workbooks.Open(fi.FullName)
    End Sub

    Public Sub CreateDocument() Implements IExcel.CreateDocument
        CheckCreated()
        If _doc IsNot Nothing Then Throw New Exception("Excel Document is already opened. Ise CloseDocument before open new one.")
        _doc = _app.Workbooks.Add
    End Sub

    Public Sub SaveDocument(Optional newPath As String = "") Implements IExcel.SaveDocument
        CheckDocument()
        If newPath > "" Then
            _doc.SaveAs2(newPath)
        Else
            _doc.Save()
        End If
    End Sub

    Public Sub CloseDocument() Implements IExcel.CloseDocument
        CheckDocument()
        _doc.Close(False)
    End Sub

    Public Sub FindAndReplace(findText As String, replaceText As String) Implements IExcel.FindAndReplace
        _doc.Activate()
        For i = 1 To _doc.Worksheets.Count
            Dim wsh As Worksheet = _doc.Worksheets(i)
            wsh.Cells.Replace(findText, replaceText)
        Next
    End Sub
End Class
