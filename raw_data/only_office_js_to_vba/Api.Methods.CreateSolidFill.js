# Create a solid fill and add a shape to the active worksheet
# Создание сплошного заполнения и добавление фигуры на активный лист

```vba
' VBA Code to create a solid fill and add a shape to the active worksheet

Sub AddShapeWithFill()
    Dim oWorksheet As Worksheet
    Dim oRGBColor As Long
    Dim oFill As FillFormat
    Dim oStroke As LineFormat
    Dim shape As Shape
    
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Create RGB color
    oRGBColor = RGB(255, 111, 61)
    
    ' Add a shape (Flowchart Online Storage) at specified positions with default size
    Set shape = oWorksheet.Shapes.AddShape(msoShapeFlowchartDatabase, _
                                          60, 35, 100, 100)
    
    ' Apply solid fill color to the shape
    With shape.Fill
        .ForeColor.RGB = oRGBColor
        .Solid
    End With
    
    ' Remove the stroke (outline) from the shape
    With shape.Line
        .Visible = msoFalse
    End With
End Sub
```

```javascript
// JavaScript Code to create a solid fill and add a shape to the active sheet

// This example creates a solid fill to apply to the object using a selected solid color as the object background.
var oWorksheet = Api.GetActiveSheet();

// Create RGB color
var oRGBColor = Api.CreateRGBColor(255, 111, 61);

// Create solid fill with the RGB color
var oFill = Api.CreateSolidFill(oRGBColor);

// Create stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add the shape to the worksheet with specified parameters
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
);
```