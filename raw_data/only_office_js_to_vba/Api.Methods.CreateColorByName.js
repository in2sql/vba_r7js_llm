# Description
**English:** This code sets the value of cell A2 to "Text with color" and changes its font color to "peachPuff".

**Russian:** Этот код устанавливает значение ячейки A2 на "Text with color" и изменяет цвет шрифта на "peachPuff".

## VBA Code
```vba
' This example sets the font color of cell A2 to PeachPuff
Sub SetCellColor()
    Dim ws As Worksheet
    Dim cell As Range
    Dim color As Long
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    ' Reference to cell A2
    Set cell = ws.Range("A2")
    ' Define the PeachPuff color using RGB
    color = RGB(255, 218, 185) ' PeachPuff color
    ' Set the value of cell A2
    cell.Value = "Text with color"
    ' Set the font color of cell A2
    cell.Font.Color = color
End Sub
```

## OnlyOffice JS Code
```javascript
// This example creates a color selecting it from one of the available color presets.
var oWorksheet = Api.GetActiveSheet();
// Create a color by the name "peachPuff"
var oColor = Api.CreateColorByName("peachPuff");
// Set the value of cell A2
oWorksheet.GetRange("A2").SetValue("Text with color");
// Set the font color of cell A2 to "peachPuff"
oWorksheet.GetRange("A2").SetFontColor(oColor);
```