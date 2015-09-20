Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Word

Public Class Word2013

    Private _word As Application

    Public Sub CreateWord()
        If _word IsNot Nothing Then Throw New Exception("Word Application already created.")
        _word = New Application
    End Sub

    Public Sub CloseWord()
        _word.Quit()
        _word = Nothing
    End Sub

    Public Shared Sub TestWorking()
        Dim word As New Application
        word.Visible = True
        word.Visible = False
        word.Quit()
    End Sub

    Public Sub Show()

    End Sub
End Class
