**Description:**
This code sets the font color of cell A2 to a specified RGB color.
Этот код устанавливает цвет шрифта ячейки A2 в указанный RGB цвет.

```vba
' VBA Code to set font color of cell A2

Sub SetFontColor()
    Dim ws As Worksheet
    Dim rgbColor As Long
    
    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Create RGB color
    rgbColor = RGB(255, 111, 61)
    
    ' Set value in cell A2
    ws.Range("A2").Value = "Text with color"
    
    ' Set font color of cell A2
    ws.Range("A2").Font.Color = rgbColor
End Sub
```

```javascript
// JavaScript Code to set font color of cell A2 using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create RGB color
var oColor = Api.CreateColorFromRGB(255, 111, 61);

// Set value in cell A2
oWorksheet.GetRange("A2").SetValue("Text with color");

// Set font color of cell A2
oWorksheet.GetRange("A2").SetFontColor(oColor);
```