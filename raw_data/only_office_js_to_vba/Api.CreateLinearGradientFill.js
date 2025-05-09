```javascript
// This code creates a linear gradient fill and applies it to a shape added to the active worksheet.
// Этот код создает линейный градиент и применяет его к фигуре, добавленной на активный лист.

// OnlyOffice JavaScript Code
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oGs1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0); // Create first gradient stop
var oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000); // Create second gradient stop
var oFill = Api.CreateLinearGradientFill([oGs1, oGs2], 5400000); // Create linear gradient fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create stroke with no fill
oWorksheet.AddShape(
    "flowChartOnlineStorage",
    60 * 36000,
    35 * 36000,
    oFill,
    oStroke,
    0,
    2 * 36000,
    1,
    3 * 36000
); // Add shape to worksheet with specified properties
```

```vba
' This code creates a linear gradient fill and applies it to a shape added to the active worksheet.
' Этот код создает линейный градиент и применяет его к фигуре, добавленной на активный лист.

' Excel VBA Code
Sub AddGradientShape()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' Get the active worksheet
    
    ' Add a rectangle shape to the worksheet
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartOnlineStorage, 60, 35, 200, 100) ' Adjust width and height as needed
    
    ' Set the fill to a linear gradient
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 213, 191) ' First gradient color
        .BackColor.RGB = RGB(255, 111, 61) ' Second gradient color
        .TwoColorGradient msoGradientHorizontal, 1 ' Create a horizontal linear gradient
    End With
    
    ' Remove the line (stroke) from the shape
    With shp.Line
        .Visible = msoFalse
    End With
End Sub
```