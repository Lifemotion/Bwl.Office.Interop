Imports System.Drawing

Module TestApp

    Sub Main()
        'Excel()
        'Word()
        Word3()
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

    Sub Word2()
        Dim appDebug As New Word2013
        Dim app As IWord = appDebug
        app.Create()
        app.CreateDocument()
        'app.OpenDocumentCopy("test.doc", DateTime.Now.Ticks.ToString() + ".doc")
        app.Visible = True
        '-----------------------------------------------------------------------
        app.AppendText("Заголовок", 24, True)
        app.AppendText()
        '------------------------------
        Dim bmp As New Bitmap(100, 100)
        app.AddPicture(bmp)
        '------------------------------
        app.AppendText("Это тестовый текст! Строка шрифтом 8", 8, True)
        app.AppendText("Это тестовый текст! Строка шрифтом 9", 9, True)
        app.AppendText("Это тестовый текст! Строка шрифтом 10", 10, True)
        app.AppendText("Это тестовый текст! Строка шрифтом 11", 11, True)
        app.AppendText()
        Dim tableIdx1 = app.AddTable(10, 4, "Новая таблица 1")
        app.SetTableText(tableIdx1, 1, 1, "1;2", False, 11)
        app.SetTableText(tableIdx1, 2, 2, "2;2", False, 11)
        app.SetTableText(tableIdx1, 3, 3, "3;3", False, 11)
        app.AppendText()
        '-----------------------------
        app.AddPicture("cat.jpg")
        '-----------------------------
        app.AppendText()
        app.AppendText("Это тестовый текст! Строка шрифтом 11", 11, True)
        app.AppendText("Это тестовый текст! Строка шрифтом 10", 10, True)
        app.AppendText("Это тестовый текст! Строка шрифтом 9", 9, True)
        app.AppendText("Это тестовый текст! Строка шрифтом 8", 8, True)
        app.AppendText()
        Dim tableIdx2 = app.AddTable(10, 4, "Новая таблица 2")
        app.SetTableText(tableIdx2, 1, 1, "1;2", False, 11)
        app.SetTableText(tableIdx2, 2, 2, "2;2", False, 11)
        app.SetTableText(tableIdx2, 3, 3, "3;3", False, 11)
        '-----------------------------------------------------------------------
        Console.ReadLine()
        app.CloseDocument()
        app.Close()
    End Sub

    Sub Word3()
        Dim orlanReport = New OrlanReportBuilder()
        orlanReport.ViolationsReport()
    End Sub

End Module
