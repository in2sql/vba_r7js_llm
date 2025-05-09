# Description / Описание
**English:** This code creates a radial gradient fill and applies it to a shape in the active worksheet.

**Русский:** Этот код создает радиальный градиент и применяет его к фигуре на активном листе.

```javascript
// JavaScript Code for OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create the first gradient stop with RGB color (255, 213, 191) at position 0
var oGs1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0);

// Create the second gradient stop with RGB color (255, 111, 61) at position 100000
var oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000);

// Create a radial gradient fill with the defined gradient stops
var oFill = Api.CreateRadialGradientFill([oGs1, oGs2]);

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape with the specified properties to the worksheet
oWorksheet.AddShape(
    "flowChartOnlineStorage",      // Shape type
    60 * 36000,                    // Left position
    35 * 36000,                    // Top position
    oFill,                         // Fill
    oStroke,                       // Stroke
    0,                             // Rotation
    2 * 36000,                     // Width
    1,                             // Height
    3 * 36000                      // Z-order
);
```

```vba
' VBA Code Equivalent

Sub CreateRadialGradientShape()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Create the first gradient stop with RGB color (255, 213, 191) at position 0%
    Dim oGs1 As GradientStop
    Set oGs1 = oWorksheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 0, 0).Fill.GradientStops.Add(0, RGB(255, 213, 191))
    
    ' Create the second gradient stop with RGB color (255, 111, 61) at position 100%
    Dim oGs2 As GradientStop
    Set oGs2 = oWorksheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 0, 0).Fill.GradientStops.Add(1, RGB(255, 111, 61))
    
    ' Create a radial gradient fill with the defined gradient stops
    With oWorksheet.Shapes.AddShape(msoShapeFlowchartStorage, 60, 35, 200, 100).Fill
        .OneColorGradient msoGradientRadial, 1, 1
        .GradientStops.Clear
        .GradientStops.Insert RGB(255, 213, 191), 0
        .GradientStops.Insert RGB(255, 111, 61), 1
    End With
    
    ' Set the stroke properties (no fill)
    With oWorksheet.Shapes(oWorksheet.Shapes.Count).Line
        .Visible = msoFalse
    End With
End Sub
```