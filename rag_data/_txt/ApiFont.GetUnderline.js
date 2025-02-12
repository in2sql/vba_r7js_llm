# Description
This code demonstrates how to set and retrieve the underline property of a specific range in a worksheet.

Этот код демонстрирует, как установить и получить свойство подчеркивания для определенного диапазона в листе.

```vba
' VBA code equivalent to the OnlyOffice JS example

Sub SetUnderlineProperty()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim oCharacters As Characters
    Dim oFont As Font
    Dim sUnderline As String

    ' Get the active worksheet
    Set oWorksheet = ActiveSheet

    ' Get range B1 and set its value
    Set oRange = oWorksheet.Range("B1")
    oRange.Value = "This is just a sample text."

    ' Get characters from position 9 with length 4
    Set oCharacters = oRange.Characters(Start:=9, Length:=4)

    ' Get the font of the characters
    Set oFont = oCharacters.Font

    ' Set underline style to single
    oFont.Underline = xlUnderlineStyleSingle

    ' Get the underline style
    sUnderline = oFont.Underline

    ' Set the value in cell B3
    oWorksheet.Range("B3").Value = "Underline property: " & sUnderline
End Sub
```

```javascript
// OnlyOffice JS code equivalent to the Excel VBA example

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get range B1 and set its value
var oRange = oWorksheet.GetRange("B1");
oRange.SetValue("This is just a sample text.");

// Get characters from position 9 with length 4
var oCharacters = oRange.GetCharacters(9, 4);

// Get the font of the characters
var oFont = oCharacters.GetFont();

// Set underline style to single
oFont.SetUnderline("xlUnderlineStyleSingle");

// Get the underline style
var sUnderline = oFont.GetUnderline();

// Set the value in cell B3
oWorksheet.GetRange("B3").SetValue("Underline property: " + sUnderline);
```