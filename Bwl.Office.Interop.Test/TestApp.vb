Module TestApp


    Sub Main()
        Excel()
        'Word()
    End Sub


    Sub Excel()
        Dim appDebug As New Excel2013
        Dim app As IExcel = appDebug

        app.Create()
        '   word.CreateDocument()
        Dim rnd As New Random
        app.OpenDocumentCopy("test.xls", rnd.Next.ToString + ".xls")

        app.Visible = True
        app.FindAndReplace("#cat2", "Test111")

        Console.ReadLine()
        app.CloseDocument()
        app.Close()
    End Sub

    Sub Word()
        Dim appDebug As New Word2013
        Dim app As IWord = appDebug

        app.Create()
        '   word.CreateDocument()
        Dim rnd As New Random
        app.OpenDocumentCopy("test.doc", rnd.Next.ToString + ".doc")

        app.Visible = True
        app.FindAndReplace("#cat2", "Test111")

        Console.ReadLine()
        app.CloseDocument()
        app.Close()
    End Sub

End Module
