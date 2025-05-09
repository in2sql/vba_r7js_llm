# Description

**English:**

This code retrieves the active worksheet, sets the value of cell B1, selects characters starting from the ninth position with a length of four, obtains the font of those characters, and sets the font to bold.

**Russian:**

Этот код получает активный лист, устанавливает значение ячейки B1, выбирает символы, начиная с девятой позиции длиной четыре символа, получает шрифт этих символов и устанавливает его жирным.

```vba
' VBA code equivalent
Sub SetBoldFont()
    Dim ws As Worksheet
    Dim rng As Range
    Dim chars As Characters
    Dim font As Font

    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Get the range B1
    Set rng = ws.Range("B1")

    ' Set the value of B1
    rng.Value = "This is just a sample text."

    ' Get characters starting from position 9 with length 4
    Set chars = rng.Characters(Start:=9, Length:=4)

    ' Get the font of the characters
    Set font = chars.Font

    ' Set the font to bold
    font.Bold = True
End Sub
```

```javascript
// JavaScript code equivalent

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get the range B1
var oRange = oWorksheet.GetRange("B1");

// Set the value of B1
oRange.SetValue("This is just a sample text.");

// Get characters starting from position 9 with length 4
var oCharacters = oRange.GetCharacters(9, 4);

// Get the font of the characters
var oFont = oCharacters.GetFont();

// Set the font to bold
oFont.SetBold(true);
```