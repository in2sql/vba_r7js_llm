# Description

**English**: This code sets the value of cell B1 and changes the color of specific characters within the text.

**Russian**: Этот код устанавливает значение ячейки B1 и изменяет цвет определенных символов в тексте.

```javascript
// This code sets the value of cell B1 and changes the color of specific characters within the text.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oRange = oWorksheet.GetRange("B1"); // Get cell B1
oRange.SetValue("This is just a sample text."); // Set the value of cell B1
var oCharacters = oRange.GetCharacters(9, 4); // Get characters starting at position 9 with length 4
var oFont = oCharacters.GetFont(); // Get the font of the selected characters
var oColor = Api.CreateColorFromRGB(255, 111, 61); // Create a color with RGB values (255,111,61)
oFont.SetColor(oColor); // Set the font color of the selected characters
```

```vba
' This code sets the value of cell B1 and changes the color of specific characters within the text.
Sub SetCellValueAndFontColor()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' Get the active worksheet
    
    With ws.Range("B1") ' Get cell B1
        .Value = "This is just a sample text." ' Set the value of cell B1
        .Characters(Start:=9, Length:=4).Font.Color = RGB(255, 111, 61) ' Set font color of specific characters
    End With
End Sub
```