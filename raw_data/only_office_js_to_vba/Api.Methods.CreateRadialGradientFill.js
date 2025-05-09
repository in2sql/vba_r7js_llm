**Description / Описание**

This code creates a shape with a radial gradient fill and no border in the active worksheet.
Этот код создаёт фигуру с радиальным градиентным заливкой и без границы на активном листе.

```vba
' VBA Code to create a shape with radial gradient fill and no border
Sub AddRadialGradientShape()
    Dim oWorksheet As Worksheet
    Dim oShape As Shape
    Dim gradient As Gradient
    Dim stops As GradientStops
    
    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Add a shape to the worksheet
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartCloud, _
        60, 35, 200, 100) ' Adjust size and position as needed
    
    ' Set the fill to a radial gradient
    With oShape.Fill
        .Visible = msoTrue
        .GradientStyle = msoGradientRadial
        .GradientVariant = 1
        .GradientDegree = 1
        .GradientStop.Clear
        
        ' Add first gradient stop
        With .GradientStop.Add(0)
            .Color.RGB = RGB(255, 213, 191)
        End With
        
        ' Add second gradient stop
        With .GradientStop.Add(100)
            .Color.RGB = RGB(255, 111, 61)
        End With
    End With
    
    ' Remove the border
    With oShape.Line
        .Visible = msoFalse
    End With
End Sub
```

```javascript
// JavaScript Code to create a shape with radial gradient fill and no border
function addRadialGradientShape() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Create gradient stops
    var oGs1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0);
    var oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000);
    
    // Create radial gradient fill
    var oFill = Api.CreateRadialGradientFill([oGs1, oGs2]);
    
    // Create no fill stroke
    var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    
    // Add shape to the worksheet
    oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000);
}
```