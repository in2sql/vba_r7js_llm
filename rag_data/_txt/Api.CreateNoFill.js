# This example creates no fill and removes the fill from the element.
# Этот пример не создает заливку и удаляет заливку из элемента.

## VBA Code
```vba
Sub AddShapeWithNoFill()
    Dim oWorksheet As Worksheet
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    Dim oShape As Shape
    ' Add a flowchart shape to the worksheet
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartManualInput, 60, 35, 200, 100)
    
    ' Set no fill for the shape
    With oShape.Fill
        .Visible = msoFalse
    End With
    
    ' Set no line for the shape
    With oShape.Line
        .Visible = msoFalse
    End With
End Sub
```

## OnlyOffice JS Code
```javascript
// This example creates no fill and removes the fill from the element.
var oWorksheet = Api.GetActiveSheet();

// Create gradient stops with specified RGB colors and positions
var oGs1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0);
var oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000);

// Create a linear gradient fill with the gradient stops and angle
var oFill = Api.CreateLinearGradientFill([oGs1, oGs2], 5400000);

// Create a stroke with width 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with the specified parameters
oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000);
```