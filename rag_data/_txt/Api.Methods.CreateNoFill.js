**Description:**

*English:*

This code creates a shape on the active worksheet with a linear gradient fill and no stroke.

*Russian:*

Этот код создает фигуру на активном листе с линейным градиентным заливом и без обводки.

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create first gradient stop with RGB color (255, 213, 191) at position 0
var oGs1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0);

// Create second gradient stop with RGB color (255, 111, 61) at position 100000
var oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000);

// Create linear gradient fill with the gradient stops and angle 5400000
var oFill = Api.CreateLinearGradientFill([oGs1, oGs2], 5400000);

// Create stroke with width 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add the shape to the worksheet with specified parameters
oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000);
```

```vba
' Get the active worksheet
Dim ws As Worksheet
Set ws = ThisWorkbook.ActiveSheet

' Add a flowchart shape to the worksheet
Dim shp As Shape
Set shp = ws.Shapes.AddShape(msoShapeFlowchartOfflineStorage, 60, 35, 200, 100) ' Adjusted units as Excel uses points

' Set the gradient fill
With shp.Fill
    .Visible = msoTrue
    .TwoColorGradient msoGradientHorizontal, 1
    .ForeColor.RGB = RGB(255, 213, 191) ' Gradient start color
    .BackColor.RGB = RGB(255, 111, 61)   ' Gradient end color
End With

' Remove the line (stroke) from the shape
With shp.Line
    .Visible = msoFalse
End With
```