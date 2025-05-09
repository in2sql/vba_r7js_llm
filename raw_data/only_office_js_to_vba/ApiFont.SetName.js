---
**Description / Описание:**

This code sets the font name of a specific range in a worksheet and displays the font name in another cell.

Этот код устанавливает имя шрифта для определенного диапазона на листе и отображает имя шрифта в другой ячейке.

**Excel VBA Code:**

```vba
' This macro sets the font name property to the specified font.

Sub SetFontName()
    Dim ws As Worksheet
    Dim rng As Range
    Dim ch As Characters
    Dim fontName As String
    
    ' Get the active sheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set value in cell B1
    ws.Range("B1").Value = "This is just a sample text."
    
    ' Get characters from position 9, length 4
    Set rng = ws.Range("B1")
    Set ch = rng.Characters(Start:=9, Length:=4)
    
    ' Set the font name
    ch.Font.Name = "Font 1"
    
    ' Get the font name
    fontName = ch.Font.Name
    
    ' Set value in cell B3
    ws.Range("B3").Value = "Font name: " & fontName
End Sub
```

**OnlyOffice JS Code:**

```javascript
// This example sets the font name property to the specified font.

var oWorksheet = Api.GetActiveSheet();
var oRange = oWorksheet.GetRange("B1");
oRange.SetValue("This is just a sample text.");
var oCharacters = oRange.GetCharacters(9, 4);
var oFont = oCharacters.GetFont();
oFont.SetName("Font 1");
var sFontName = oFont.GetName();
oWorksheet.GetRange("B3").SetValue("Font name: " + sFontName);
```