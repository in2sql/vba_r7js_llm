# Description / Описание

**English:**
This code sets the background color of cells A1:C1 to a specific RGB color and sets a value in cell A3 indicating that the color has been applied.

**Russian:**
Этот код устанавливает цвет фона ячеек A1:C1 на определенный RGB цвет и задает значение в ячейке A3, указывающее, что цвет был применен.

```vba
' VBA code to set the background color of range A1:C1 and update cell A3
Sub SetRangeColor()
    Dim ws As Worksheet
    Dim rng As Range
    
    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Get the range A1:C1
    Set rng = ws.Range("A1:C1")
    
    ' Set the fill color to RGB(255, 213, 191)
    rng.Interior.Color = RGB(255, 213, 191)
    
    ' Set the value in cell A3
    ws.Range("A3").Value = "The color was set to the background of cells A1:C1."
End Sub
```

```javascript
// JavaScript code to set the background color of range A1:C1 and update cell A3
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oRange = Api.GetRange("A1:C1"); // Get the range A1:C1
// Set the fill color to RGB(255, 213, 191)
oRange.SetFillColor(Api.CreateColorFromRGB(255, 213, 191));
// Set the value in cell A3
oWorksheet.GetRange("A3").SetValue("The color was set to the background of cells A1:C1.");
```