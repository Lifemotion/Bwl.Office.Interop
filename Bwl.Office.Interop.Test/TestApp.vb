Imports System.Drawing

Module TestApp

    Sub Main()
        'Excel(Of ExcelDynamic)()
        'Word()
        Word2(Of WordDynamic)()
    End Sub

    Sub Excel(Of T As {IExcel, New})()
        Dim app = New T

        app.Create()
        Dim rnd As New Random
        app.OpenDocumentCopy("test.xls", rnd.Next.ToString + ".xls")

        app.Visible = True
        app.FindAndReplace("#cat2", "Test111")

        Console.ReadLine()
        app.CloseDocument()
        app.Close()
    End Sub

    Sub Word(Of T As {IWord, New})()
        Dim appDebug As New T
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

    Sub Word2(Of T As {IWord, New})()
        Dim appDebug As New T
        Dim app As IWord = appDebug
        app.Create()
        app.CreateDocument()
        'app.OpenDocumentCopy("test.doc", DateTime.Now.Ticks.ToString() + ".doc")
        app.Visible = True
        '-----------------------------------------------------------------------
        Dim isBold = False : Dim newParagraph = True
        app.AppendText("Заголовок", New TextStyle() With {.FontName = "Arial", .FontSize = 24}, True)
        app.AppendText()
        '------------------------------
        Dim bmp As New Bitmap(100, 100)
        app.AddPicture(bmp)
        '------------------------------        
        app.AppendText("Это тестовый текст Verdana! Строка шрифтом 8", New TextStyle() With {.FontSize = 8}, newParagraph)
        app.AppendText("Это тестовый текст Verdana! Строка шрифтом 9", New TextStyle() With {.FontSize = 9}, newParagraph)
        app.AppendText("Это тестовый текст Verdana! Строка шрифтом 10", New TextStyle() With {.FontSize = 10}, newParagraph)
        app.AppendText("Это тестовый текст Verdana! Строка шрифтом 11", New TextStyle() With {.FontSize = 11}, newParagraph)
        app.AppendText()
        Dim tableIdx1 = app.AddTable(10, 4, "Новая таблица 1")
        app.SetTableText(tableIdx1, 1, 1, "1;2", New TextStyle())
        app.SetTableText(tableIdx1, 2, 2, "2;2", New TextStyle())
        app.SetTableText(tableIdx1, 3, 3, "3;3", New TextStyle())
        app.AppendText()
        '-----------------------------
        app.AddPicture("cat.jpg")
        '-----------------------------
        app.AppendText()
        app.AppendText("Это тестовый текст Times! Строка шрифтом 8", New TextStyle() With {.FontName = "Times New Roman", .FontSize = 8}, newParagraph)
        app.AppendText("Это тестовый текст Times! Строка шрифтом 9", New TextStyle() With {.FontName = "Times New Roman", .FontSize = 9}, newParagraph)
        app.AppendText("Это тестовый текст Times! Строка шрифтом 10", New TextStyle() With {.FontName = "Times New Roman", .FontSize = 10}, newParagraph)
        app.AppendText("Это тестовый текст Times! Строка шрифтом 11", New TextStyle() With {.FontName = "Times New Roman", .FontSize = 11}, newParagraph)
        app.AppendText()
        Dim tableIdx2 = app.AddTable(10, 4, "Новая таблица 2")
        app.SetTableText(tableIdx2, 1, 1, "1;2", New TextStyle() With {.IsBold = True})
        app.SetTableText(tableIdx2, 2, 2, "2;2", New TextStyle() With {.IsBold = True})
        app.SetTableText(tableIdx2, 3, 3, "3;3", New TextStyle() With {.IsBold = True})
        '-----------------------------------------------------------------------
        Console.ReadLine()
        app.CloseDocument()
        app.Close()
    End Sub
End Module
