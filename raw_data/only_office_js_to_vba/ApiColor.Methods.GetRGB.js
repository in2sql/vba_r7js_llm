```plaintext
// Sets the value and font color of cell A2 and displays the RGB color in cell A4
// Устанавливает значение и цвет шрифта ячейки A2 и отображает цвет RGB в ячейке A4
```

```vba
' VBA code to set cell values and font color

Sub SetCellValuesAndColors()
    Dim oWorksheet As Worksheet
    Dim oColor As Long
    Dim nColor As Long
    
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Create a color from RGB values
    oColor = RGB(255, 111, 61)
    
    ' Set value of cell A2
    oWorksheet.Range("A2").Value = "Text with color"
    
    ' Set font color of cell A2
    oWorksheet.Range("A2").Font.Color = oColor
    
    ' Get the RGB value of the color
    nColor = oColor
    
    ' Set value of cell A4 with RGB color
    oWorksheet.Range("A4").Value = "Cell color in RGB format: " & nColor
End Sub
```

```javascript
// JavaScript code to set cell values and font color using OnlyOffice API

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oColor = Api.CreateColorFromRGB(255, 111, 61); // Create color from RGB values
oWorksheet.GetRange("A2").SetValue("Text with color"); // Set value of cell A2
oWorksheet.GetRange("A2").SetFontColor(oColor); // Set font color of cell A2
var nColor = oColor.GetRGB(); // Get RGB value of the color
oWorksheet.GetRange("A4").SetValue("Cell color in RGB format: " + nColor); // Set value of cell A4 with RGB color
```