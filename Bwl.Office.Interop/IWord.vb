Imports System.Drawing

Public Enum TableStyle
    wdStyleTableColorfulGrid
    wdStyleTableColorfulList
    wdStyleTableColorfulShading
    wdStyleTableDarkList
    wdStyleTableLightGrid
    wdStyleTableLightGridAccent1
    wdStyleTableLightList
    wdStyleTableLightListAccent1
    wdStyleTableLightShading
    wdStyleTableLightShadingAccent1
    wdStyleTableMediumGrid1
    wdStyleTableMediumGrid2
    wdStyleTableMediumGrid3
    wdStyleTableMediumList1
    wdStyleTableMediumList1Accent1
    wdStyleTableMediumList2
    wdStyleTableMediumShading1
    wdStyleTableMediumShading1Accent1
    wdStyleTableMediumShading2
    wdStyleTableMediumShading2Accent1
End Enum

Public Interface IWord
    Inherits IOfficeApp
    Sub AppendText(Optional fontSize As Integer = 6)
    Sub AppendText(text As String, fontSize As Integer, newParagraph As Boolean)
    Function AddTable(nRows As Integer, nCols As Integer, caption As String, Optional style As TableStyle = TableStyle.wdStyleTableLightGrid) As Integer
    Sub SetTableText(tableIdx As Integer, row As Integer, col As Integer, text As String, IsBold As Boolean, fontSize As Integer)
    Sub AddPicture(fileName As String)
    Sub AddPicture(bitmap As Bitmap)
End Interface
