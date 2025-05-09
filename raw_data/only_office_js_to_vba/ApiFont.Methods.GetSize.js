```javascript
// This code sets a value in cell B1, modifies the font size of specific characters, and writes the font size to cell B3.
// Этот код устанавливает значение в ячейке B1, изменяет размер шрифта определенных символов и записывает размер шрифта в ячейку B3.

var oWorksheet = Api.GetActiveSheet();
var oRange = oWorksheet.GetRange("B1");
oRange.SetValue("This is just a sample text."); // Set value in B1
var oCharacters = oRange.GetCharacters(9, 4); // Get characters from position 9 with length 4
var oFont = oCharacters.GetFont(); // Get the font of the selected characters
oFont.SetSize(18); // Set font size to 18
var nSize = oFont.GetSize(); // Get the current font size
oWorksheet.GetRange("B3").SetValue("Size property: " + nSize); // Write font size to B3
```

```vba
' This code sets a value in cell B1, modifies the font size of specific characters, and writes the font size to cell B3.
' Этот код устанавливает значение в ячейке B1, изменяет размер шрифта определенных символов и записывает размер шрифта в ячейку B3.

Sub ModifyFontSize()
    Dim ws As Worksheet
    Dim rng As Range
    Dim ch As Characters
    Dim fontSize As Double
    
    Set ws = ThisWorkbook.ActiveSheet
    Set rng = ws.Range("B1")
    rng.Value = "This is just a sample text." ' Set value in B1
    
    ' Modify characters from position 9, length 4
    Set ch = rng.Characters(Start:=9, Length:=4)
    ch.Font.Size = 18 ' Set font size to 18
    
    ' Get the font size
    fontSize = ch.Font.Size
    
    ' Write the font size to B3
    ws.Range("B3").Value = "Size property: " & fontSize
End Sub
```