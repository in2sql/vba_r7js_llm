VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GStdFonts"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Type MYFont
    Bold As Boolean
    Italic As Boolean
    Charset As Integer
    name As String
    Size As Currency
    Strikethrough As Boolean
    Underline As Boolean
End Type

Public Sub CopytoIfont(srcfont As MYFont, dstFont As IFontDisp)

    With dstFont
        .Bold = srcfont.Bold
        .Italic = srcfont.Italic
        .name = srcfont.name
        .Size = srcfont.Size
        .Strikethrough = srcfont.Strikethrough
        .Underline = srcfont.Underline
        .Charset = srcfont.Charset
    End With

End Sub

Public Sub Setfont(dstobject As Object, srcfont As MYFont)

    With dstobject
        .FontBold = srcfont.Bold
        .FontItalic = srcfont.Italic

        If srcfont.name = "" Then .FontName = "System" Else .FontName = srcfont.name
        .FontSize = srcfont.Size
        .FontStrikethru = srcfont.Strikethrough
        .FontUnderline = srcfont.Underline
    End With

End Sub
Public Sub FontEqual(dstFont As StdFont, srcfont As StdFont)
    On Error Resume Next
    With dstFont
            .Size = srcfont.Size
            .Bold = srcfont.Bold
            .Charset = srcfont.Charset
            .Italic = srcfont.Italic
            .name = srcfont.name
            .Strikethrough = srcfont.Strikethrough
            .Underline = srcfont.Underline
            .Weight = srcfont.Weight
    End With

End Sub

Public Function toMYfont(srcfont As IFontDisp) As MYFont

    With toMYfont
        .Bold = srcfont.Bold
        .Italic = srcfont.Italic
        .name = srcfont.name
        .Size = srcfont.Size
        .Strikethrough = srcfont.Strikethrough
        .Underline = srcfont.Underline
        .Charset = srcfont.Charset
    End With

End Function


