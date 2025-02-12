## Description / Описание

Creates a flowchart shape with a linear gradient fill on the active worksheet.  
Создает фигуру блок-схемы с линейным градиентным заливом на активном листе.

```javascript
// JavaScript Code using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create the first gradient stop with RGB color (255, 213, 191) at position 0
var oGs1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0);

// Create the second gradient stop with RGB color (255, 111, 61) at position 100000
var oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000);

// Create a linear gradient fill with the two gradient stops and a rotation of 5400000
var oFill = Api.CreateLinearGradientFill([oGs1, oGs2], 5400000);

// Create a stroke with weight 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with the specified parameters
oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000); 
```

```vba
' VBA Code Equivalent

Sub AddFlowChartShape()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Define RGB colors for gradient stops
    Dim color1 As Long
    color1 = RGB(255, 213, 191) ' First color
    
    Dim color2 As Long
    color2 = RGB(255, 111, 61) ' Second color
    
    ' Add a shape to the worksheet
    Dim shp As Shape
    Set shp = oWorksheet.Shapes.AddShape(msoShapeFlowchartProcess, _
                                        60, 35, _ ' Left and Top positions
                                        200, 100)   ' Width and Height
                                        
    ' Apply gradient fill to the shape
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = color1
        .BackColor.RGB = color2
        .TwoColorGradient msoGradientHorizontal, 1
        .GradientAngle = 540 ' VBA uses degrees for angle
    End With
    
    ' Remove the stroke (outline) from the shape
    With shp.Line
        .Visible = msoFalse
    End With
End Sub
```