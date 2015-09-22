Imports Bwl.Office.Interop

Public Class OrlanReportBuilder

    Dim _app As IWord = New Word2013()

    Public Sub ViolationsReport()
        CreateNewDocument()
        _app.AppendText("Отчет по нарушениям Orlan", 24, True)
        _app.AppendText(DateTime.Now.ToString("dd.MM.yyyy HH:mm"), 11, True)
        _app.AppendText()
        _app.AddTable(10, 5, "нарушение1")
    End Sub

    Private Sub CreateNewDocument()
        _app.Create()
        _app.CreateDocument()
        _app.Visible = True
    End Sub
End Class
