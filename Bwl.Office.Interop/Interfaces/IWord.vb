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
    wdStyleNormal
    wdStyleHeading4
    wdStyleHeading3
    wdStyleHeading2
    wdStyleHeading1
    wdStyleIndex4
    wdStyleIndex3
    wdStyleIndex2
    wdStyleIndex1
End Enum

Public Class TextStyle
    Public Property FontName As String = "Verdana"
    Public Property FontSize As Integer = 11
    Public Property IsBold As Boolean
    Public Property IsItalic As Boolean
End Class

Public Interface IWord
    Inherits IOfficeApp
    Sub AppendText()
    Sub AppendText(text As String, style As TextStyle, newParagraph As Boolean)
    Function AddTable(nRows As Integer, nCols As Integer, caption As String, Optional style As TableStyle = TableStyle.wdStyleTableLightGrid) As Integer
    Sub SetTableText(tableIdx As Integer, row As Integer, col As Integer, text As String, style As TextStyle)
    Sub AddPicture(fileName As String)
    Sub AddPicture(bitmap As Bitmap)
End Interface
