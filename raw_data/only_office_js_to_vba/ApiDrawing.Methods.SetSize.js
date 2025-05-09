# Example of setting the size and position of a shape in a worksheet.
# Пример установки размера и позиции фигуры на листе.

```javascript
// This example sets the size of the shape bounding box.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill with specified RGB color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
// Add a shape with specified type, position, fill, stroke, and other parameters
var oDrawing = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
oDrawing.SetSize(120 * 36000, 70 * 36000); // Set the size of the shape
oDrawing.SetPosition(0, 2 * 36000, 2, 3 * 36000); // Set the position of the shape
```

```vba
' This example sets the size and position of a shape on the active worksheet
Sub SetShapeSizeAndPosition()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' Get the active worksheet
    
    Dim shp As Shape
    ' Add a shape with specified type and position (Left, Top, Width, Height)
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartBlockEnd, 60, 35, 120, 70)
    
    ' Set the fill color using RGB values
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61)
        .Solid
    End With
    
    ' Remove the stroke
    With shp.Line
        .Visible = msoFalse
    End With
    
    ' Set the position of the shape
    shp.Left = 60 ' Left position in points
    shp.Top = 35 ' Top position in points
End Sub
```