# Description
**English:** This code creates a color from a preset, sets the value of cell A2, and applies the font color to the cell.

**Russian:** Этот код создает цвет из предустановленного набора, устанавливает значение ячейки A2 и применяет цвет шрифта к ячейке.

```vba
' VBA Code to create a color, set cell A2 value, and apply font color

Sub SetCellValueAndColor()
    Dim ws As Worksheet
    Dim cell As Range
    Dim peachPuffColor As Long
    
    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Define the color PeachPuff using RGB
    peachPuffColor = RGB(255, 218, 185)
    
    ' Set the value of cell A2
    Set cell = ws.Range("A2")
    cell.Value = "Text with color"
    
    ' Apply the font color to cell A2
    cell.Font.Color = peachPuffColor
End Sub
```

```javascript
// JavaScript Code to create a color, set cell A2 value, and apply font color

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a color by name "peachPuff"
var oColor = Api.CreateColorByName("peachPuff");

// Set the value of cell A2
oWorksheet.GetRange("A2").SetValue("Text with color");

// Apply the font color to cell A2
oWorksheet.GetRange("A2").SetFontColor(oColor);
```