# Set strikethrough for specific characters in a cell / Установка зачёркивания для определённых символов в ячейке

Sets the strikethrough property to characters 9 through 12 in cell B1 containing "This is just a sample text."
Устанавливает свойство зачёркивания для символов с 9 по 12 в ячейке B1, содержащей текст "This is just a sample text."

```vba
' VBA Code to set strikethrough for specific characters in cell B1

Sub SetStrikethrough()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range

    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Set the range to cell B1
    Set rng = ws.Range("B1")

    ' Set the cell value
    rng.Value = "This is just a sample text."

    ' Apply strikethrough to characters 9 to 12
    With rng.Characters(Start:=9, Length:=4).Font
        .Strikethrough = True
    End With
End Sub
```

```javascript
// JavaScript Code to set strikethrough for specific characters in cell B1

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get the range for cell B1
var oRange = oWorksheet.GetRange("B1");

// Set the value of cell B1
oRange.SetValue("This is just a sample text.");

// Get characters from position 9 with length 4
var oCharacters = oRange.GetCharacters(9, 4);

// Get the font of the specified characters
var oFont = oCharacters.GetFont();

// Set the strikethrough property to true
oFont.SetStrikethrough(true);
```