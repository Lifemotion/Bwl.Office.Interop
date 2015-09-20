Imports Bwl.Office.Interop
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Word

Public Class Word2013
    Implements IWord

    Private _app As Application
    Private _doc As Document

    Public Property Visible As Boolean Implements IWord.Visible
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
        Dim word As New Application
        word.Visible = True
        word.Visible = False
        word.Quit()
    End Sub

    Private Sub CheckCreated()
        If _app Is Nothing Then Throw New Exception("Word Application not created. Use Create method before.")
    End Sub

    Private Sub CheckDocument()
        CheckCreated()
        If _doc Is Nothing Then Throw New Exception("Word Document not opened. Use OpenDocument method before.")
    End Sub

    Public Sub Create() Implements IWord.Create
        If _app IsNot Nothing Then Throw New Exception("Word Application already created.")
        _app = New Application
    End Sub

    Public Sub Close() Implements IWord.Close
        _app.Quit()
        _app = Nothing
    End Sub

    Public Sub OpenDocument(path As String) Implements IWord.OpenDocument
        CheckCreated()
        Dim fi As New IO.FileInfo(path)
        If _doc IsNot Nothing Then Throw New Exception("Document is already opened. Ise CloseDocument before open new one.")
        _doc = _app.Documents.Open(fi.FullName)
    End Sub

    Public Sub OpenDocumentCopy(originalPath As String, copyPath As String) Implements IWord.OpenDocumentCopy
        CheckCreated()
        Dim fi As New IO.FileInfo(originalPath)
        fi.CopyTo(copyPath)
        fi = New IO.FileInfo(copyPath)
        If _doc IsNot Nothing Then Throw New Exception("Document is already opened. Ise CloseDocument before open new one.")
        _doc = _app.Documents.Open(fi.FullName)
    End Sub

    Public Sub CreateDocument() Implements IWord.CreateDocument
        CheckCreated()
        If _doc IsNot Nothing Then Throw New Exception("Document is already opened. Ise CloseDocument before open new one.")
        _doc = _app.Documents.Add
    End Sub

    Public Sub SaveDocument(Optional newPath As String = "") Implements IWord.SaveDocument
        CheckDocument()
        If newPath > "" Then
            _doc.SaveAs2(newPath)
        Else
            _doc.Save()
        End If
    End Sub

    Public Sub CloseDocument() Implements IWord.CloseDocument
        CheckDocument()
        _doc.Close(False)
    End Sub

    Public Sub FindAndReplace(findText As String, replaceText As String) Implements IWord.FindAndReplace
        _doc.Activate()
        _app.Selection.Find.Execute(findText, False, False, False, False, False, True, 1, False, replaceText, 2, False, False, False, False)
    End Sub
End Class
