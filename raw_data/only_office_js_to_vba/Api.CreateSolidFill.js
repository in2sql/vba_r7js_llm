# Create a Solid Fill and Add a Shape to the Active Worksheet
# Создание сплошной заливки и добавление фигуры на активный лист

```vba
' VBA code to create a solid fill and add a shape to the active worksheet

Sub AddShapeWithSolidFill()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Define RGB color
    Dim oRGBColor As Long
    oRGBColor = RGB(255, 111, 61)
    
    ' Add a shape with solid fill
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, _
        60 * 72, 35 * 72, 200, 100) ' Position and size in points (1 inch = 72 points)
    
    ' Apply solid fill color
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = oRGBColor
        .Solid
    End With
    
    ' Remove stroke
    With oShape.Line
        .Visible = msoFalse
    End With
End Sub
```

```javascript
// JavaScript code to create a solid fill and add a shape to the active worksheet

function addShapeWithSolidFill() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Define RGB color
    var oRGBColor = Api.CreateRGBColor(255, 111, 61);
    
    // Create a solid fill with the specified color
    var oFill = Api.CreateSolidFill(oRGBColor);
    
    // Remove stroke by setting it to no fill
    var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    
    // Add the shape to the worksheet with specified properties
    oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000);
}
```