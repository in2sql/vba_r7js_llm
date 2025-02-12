**Description / Описание**

This code example creates a shape with a linear gradient fill using preset and custom colors in Both JavaScript and Excel VBA.

Этот пример кода создает фигуру с линейным градиентным заполнением, используя предустановленные и пользовательские цвета как в JavaScript, так и в Excel VBA.

```vba
' VBA Code to create a shape with a linear gradient fill

Sub AddGradientShape()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Define preset color "peachPuff"
    ' VBA does not have preset colors by name, so we use RGB equivalent
    Dim peachPuff As Long
    peachPuff = RGB(255, 218, 185)
    
    ' Create gradient stops
    Dim gradientStops(1) As GradientStop
    Set gradientStops(0) = oWorksheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 100, 100).Fill.GradientStops.Insert(1)
    gradientStops(0).Color = peachPuff
    gradientStops(0).Position = 0
    
    Set gradientStops(1) = oWorksheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 100, 100).Fill.GradientStops.Insert(2)
    gradientStops(1).Color = RGB(255, 111, 61)
    gradientStops(1).Position = 1
    
    ' Add a shape with gradient fill
    Dim shp As Shape
    Set shp = oWorksheet.Shapes.AddShape(msoShapeFlowchartStorage, 60, 35, 200, 150)
    
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = peachPuff
        .BackColor.RGB = RGB(255, 111, 61)
        .TwoColorGradient msoGradientHorizontal, 1
    End With
    
    ' Remove the stroke
    With shp.Line
        .Visible = msoFalse
    End With
End Sub
```

```javascript
// JavaScript Code to create a shape with a linear gradient fill using OnlyOffice API

// This example creates a shape with a linear gradient fill using preset and custom colors.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oPresetColor = Api.CreatePresetColor("peachPuff"); // Create a preset color "peachPuff"
var oGs1 = Api.CreateGradientStop(oPresetColor, 0); // Create first gradient stop with preset color at position 0
var oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 1); // Create second gradient stop with custom RGB color at position 1
var oFill = Api.CreateLinearGradientFill([oGs1, oGs2], 0); // Create linear gradient fill with the gradient stops
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create no stroke
oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000); // Add the shape to the worksheet
```