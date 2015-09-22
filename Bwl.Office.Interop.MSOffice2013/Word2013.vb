Imports System.IO
Imports System.Drawing
Imports System.Drawing.Imaging
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Word
Imports Bwl.Office.Interop

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
        If _doc IsNot Nothing Then Throw New Exception("Document is already opened. Use CloseDocument before open new one.")
        _doc = _app.Documents.Open(fi.FullName)
    End Sub

    Public Sub OpenDocumentCopy(originalPath As String, copyPath As String) Implements IWord.OpenDocumentCopy
        CheckCreated()
        Dim fi As New IO.FileInfo(originalPath)
        fi.CopyTo(copyPath)
        fi = New IO.FileInfo(copyPath)
        If _doc IsNot Nothing Then Throw New Exception("Document is already opened. Use CloseDocument before open new one.")
        _doc = _app.Documents.Open(fi.FullName)
    End Sub

    Public Sub CreateDocument() Implements IWord.CreateDocument
        CheckCreated()
        If _doc IsNot Nothing Then Throw New Exception("Document is already opened. Use CloseDocument before open new one.")
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

    Public Sub AppendText(Optional fontSize As Integer = 6) Implements IWord.AppendText
        _doc.Activate()
        With _app.Selection
            .Start = _doc.Range.End
            .End = _doc.Range.End
            .Font.Size = fontSize
            .TypeParagraph()
        End With
    End Sub

    Public Sub AppendText(text As String, fontSize As Integer, newParagraph As Boolean) Implements IWord.AppendText
        _doc.Activate()
        With _app.Selection
            .Start = _doc.Range.End
            .End = _doc.Range.End
            .Font.Size = fontSize
            .TypeText(text)
        End With
        If newParagraph Then _app.Selection.TypeParagraph()
    End Sub

    Public Function AddTable(nRows As Integer, nCols As Integer, caption As String, Optional style As TableStyle = TableStyle.wdStyleTableLightGrid) As Integer Implements IWord.AddTable
        _doc.Activate()
        Dim docRange As Range = _doc.Range()
        SetRange(docRange, docRange.End, docRange.End, caption)
        Dim tbl = _doc.Tables.Add(docRange, nRows, nCols)
        tbl.Style = GetWdBuiltinStyle(style)
        AppendText()
        '-----------------------
        Return _doc.Tables.Count
    End Function

    Public Sub SetTableText(tableIdx As Integer, row As Integer, col As Integer, text As String, IsBold As Boolean, fontSize As Integer) Implements IWord.SetTableText
        With _doc.Tables(tableIdx).Range
            .Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter
            .Font.Size = fontSize
            .Font.Bold = If(IsBold, 1, 0)
        End With
        _doc.Tables(tableIdx).Cell(row, col).Range.Text = text
    End Sub

    Public Sub AddPicture(fileName As String) Implements IWord.AddPicture
        _doc.InlineShapes.AddPicture(Path.GetFullPath(fileName))
        AppendText()
    End Sub

    Public Sub AddPicture(bitmap As Bitmap) Implements IWord.AddPicture
        Dim tempFilename = Guid.NewGuid().ToString()
        Dim image = CType(bitmap, Image)
        image.Save(tempFilename, ImageFormat.Jpeg)
        AddPicture(tempFilename)
        File.Delete(tempFilename)
    End Sub

    Private Sub SetRange(rng As Range, startRange As Integer, endRange As Integer, title As String,
                         Optional fontName As String = "Verdana", Optional fontSize As Integer = 11)
        With rng
            .Start = startRange
            .End = endRange
            .InsertBefore(title)
            .Font.Name = fontName
            .Font.Size = fontSize
            .InsertParagraphAfter()
            .SetRange(rng.End, rng.End)
        End With
    End Sub

    Private Function GetWdBuiltinStyle(style As TableStyle) As WdBuiltinStyle
        Select Case style
            Case TableStyle.wdStyleTableColorfulGrid
                Return WdBuiltinStyle.wdStyleTableColorfulGrid
            Case TableStyle.wdStyleTableColorfulList
                Return WdBuiltinStyle.wdStyleTableColorfulList
            Case TableStyle.wdStyleTableColorfulShading
                Return WdBuiltinStyle.wdStyleTableColorfulShading
            Case TableStyle.wdStyleTableDarkList
                Return WdBuiltinStyle.wdStyleTableDarkList
            Case TableStyle.wdStyleTableLightGrid
                Return WdBuiltinStyle.wdStyleTableLightGrid
            Case TableStyle.wdStyleTableLightGridAccent1
                Return WdBuiltinStyle.wdStyleTableLightGridAccent1
            Case TableStyle.wdStyleTableLightList
                Return WdBuiltinStyle.wdStyleTableLightList
            Case TableStyle.wdStyleTableLightListAccent1
                Return WdBuiltinStyle.wdStyleTableLightListAccent1
            Case TableStyle.wdStyleTableLightShading
                Return WdBuiltinStyle.wdStyleTableLightShading
            Case TableStyle.wdStyleTableLightShadingAccent1
                Return WdBuiltinStyle.wdStyleTableLightShadingAccent1
            Case TableStyle.wdStyleTableMediumGrid1
                Return WdBuiltinStyle.wdStyleTableMediumGrid1
            Case TableStyle.wdStyleTableMediumGrid2
                Return WdBuiltinStyle.wdStyleTableMediumGrid2
            Case TableStyle.wdStyleTableMediumGrid3
                Return WdBuiltinStyle.wdStyleTableMediumGrid3
            Case TableStyle.wdStyleTableMediumList1
                Return WdBuiltinStyle.wdStyleTableMediumList1
            Case TableStyle.wdStyleTableMediumList1Accent1
                Return WdBuiltinStyle.wdStyleTableMediumList1Accent1
            Case TableStyle.wdStyleTableMediumList2
                Return WdBuiltinStyle.wdStyleTableMediumList2
            Case TableStyle.wdStyleTableMediumShading1
                Return WdBuiltinStyle.wdStyleTableMediumShading1
            Case TableStyle.wdStyleTableMediumShading1Accent1
                Return WdBuiltinStyle.wdStyleTableMediumShading1Accent1
            Case TableStyle.wdStyleTableMediumShading2
                Return WdBuiltinStyle.wdStyleTableMediumShading2
            Case TableStyle.wdStyleTableMediumShading2Accent1
                Return WdBuiltinStyle.wdStyleTableMediumShading2Accent1
        End Select
        '-------------
        Return Nothing
    End Function
End Class
