Imports System.IO
Imports System.Drawing
Imports System.Drawing.Imaging
Imports Bwl.Office.Interop

Public Class WordDynamic
    Implements IWord

    Private _app As Object
    Private _doc As Object

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

    Public Shared Function GetAppType() As Type
        Return Type.GetTypeFromProgID("Word.Application")
    End Function

    Public Shared Sub TestWorking()

        Dim word = Activator.CreateInstance(GetAppType)

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
        _app = Activator.CreateInstance(GetAppType)
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
    Public Sub AppendText() Implements IWord.AppendText
        _doc.Activate()
        Dim style As New TextStyle()
        With _app.Selection
            .Start = _doc.Range.End
            .End = _doc.Range.End
            .Font.Name = style.FontName
            .Font.Size = style.FontSize
            .Font.Bold = If(style.IsBold, 1, 0)
            .Font.Italic = If(style.IsItalic, 1, 0)
            .TypeParagraph()
        End With
    End Sub

    Public Sub AppendText(text As String, style As TextStyle, newParagraph As Boolean) Implements IWord.AppendText
        _doc.Activate()
        With _app.Selection
            .Start = _doc.Range.End
            .End = _doc.Range.End
            .Font.Name = style.FontName
            .Font.Size = style.FontSize
            .Font.Bold = If(style.IsBold, 1, 0)
            .Font.Italic = If(style.IsItalic, 1, 0)
            .TypeText(text)
        End With
        If newParagraph Then _app.Selection.TypeParagraph()
    End Sub

    Public Function AddTable(nRows As Integer, nCols As Integer, caption As String, Optional style As TableStyle = TableStyle.wdStyleTableLightGrid) As Integer Implements IWord.AddTable
        _doc.Activate()
        Dim docRange As Object = _doc.Range()
        SetRange(docRange, docRange.End, docRange.End, caption)
        Dim tbl = _doc.Tables.Add(docRange, nRows, nCols)
        tbl.Style = GetWdBuiltinStyle(style)
        AppendText()
        '-----------------------
        Return _doc.Tables.Count
    End Function

    Public Sub SetTableText(tableIdx As Integer, row As Integer, col As Integer, text As String, style As TextStyle) Implements IWord.SetTableText
        With _doc.Tables(tableIdx).Range
            .Cells.VerticalAlignment = 1
            .Font.Name = style.FontName
            .Font.Size = style.FontSize
            .Font.Bold = If(style.IsBold, 1, 0)
            .Font.Italic = If(style.IsItalic, 1, 0)
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

    Private Sub SetRange(rng As Object, startRange As Integer, endRange As Integer, title As String,
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

    Private Function GetWdBuiltinStyle(style As TableStyle) As Integer
        Select Case style
            Case TableStyle.wdStyleTableColorfulGrid
                Return -172
            Case TableStyle.wdStyleTableColorfulList
                Return -171
            Case TableStyle.wdStyleTableColorfulShading
                Return -170
            Case TableStyle.wdStyleTableDarkList
                Return -169
            Case TableStyle.wdStyleTableLightGrid
                Return -161
            Case TableStyle.wdStyleTableLightGridAccent1
                Return -175
            Case TableStyle.wdStyleTableLightList
                Return -160
            Case TableStyle.wdStyleTableLightListAccent1
                Return -174
            Case TableStyle.wdStyleTableLightShading
                Return -159
            Case TableStyle.wdStyleTableLightShadingAccent1
                Return -173
            Case TableStyle.wdStyleTableMediumGrid1
                Return -166
            Case TableStyle.wdStyleTableMediumGrid2
                Return -167
            Case TableStyle.wdStyleTableMediumGrid3
                Return -168
            Case TableStyle.wdStyleTableMediumList1
                Return -164
            Case TableStyle.wdStyleTableMediumList1Accent1
                Return -178
            Case TableStyle.wdStyleTableMediumList2
                Return -165
            Case TableStyle.wdStyleTableMediumShading1
                Return -162
            Case TableStyle.wdStyleTableMediumShading1Accent1
                Return -176
            Case TableStyle.wdStyleTableMediumShading2
                Return -163
            Case TableStyle.wdStyleTableMediumShading2Accent1
                Return -177
            Case TableStyle.wdStyleNormal
                Return -1
            Case TableStyle.wdStyleHeading1
                Return -2
            Case TableStyle.wdStyleIndex1
                Return -11
        End Select
        Return -1

    End Function

End Class
