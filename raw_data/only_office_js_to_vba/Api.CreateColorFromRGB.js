### Description
**English:** This example sets the value of cell A2 to "Text with color" and changes its font color to a specific RGB color.

**Russian:** Этот пример устанавливает значение ячейки A2 на "Text with color" и изменяет цвет шрифта на определенный цвет RGB.

```javascript
// This example sets the value of cell A2 to "Text with color" and changes its font color to a specific RGB color.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oColor = Api.CreateColorFromRGB(255, 111, 61); // Create an RGB color
oWorksheet.GetRange("A2").SetValue("Text with color"); // Set the value of A2
oWorksheet.GetRange("A2").SetFontColor(oColor); // Set the font color of A2
```

```vba
' This example sets the value of cell A2 to "Text with color" and changes its font color to a specific RGB color.
Sub SetCellColor()
    ' Set the value of cell A2
    Range("A2").Value = "Text with color"
    ' Set the font color of cell A2 to RGB(255, 111, 61)
    Range("A2").Font.Color = RGB(255, 111, 61)
End Sub
```