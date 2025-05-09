**Description / Описание**

English: This code adds a "flowChartOnlineStorage" shape to the active worksheet with a linear gradient fill from light orange to dark orange and no stroke.

Russian: Этот код добавляет фигуру "flowChartOnlineStorage" на активный лист с линейным градиентным заполнением от светло-оранжевого до темно-оранжевого и без обводки.

```vba
' VBA code that adds a shape with a linear gradient fill and no stroke

Sub AddFlowChartOnlineStorageShape()
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet ' Get the active worksheet
    
    ' Define gradient colors
    Dim color1 As Long
    color1 = RGB(255, 213, 191) ' Light orange
    
    Dim color2 As Long
    color2 = RGB(255, 111, 61) ' Dark orange
    
    ' Add the shape to the worksheet
    Dim shp As Shape
    ' Parameters: Type, Left, Top, Width, Height
    Set shp = oWorksheet.Shapes.AddShape(msoShapeFlowchartStorage, 60, 35, 200, 100)
    
    ' Apply linear gradient fill
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = color1
        .BackColor.RGB = color2
        .TwoColorGradient msoGradientHorizontal, 1 ' Apply two-color horizontal gradient
    End With
    
    ' Remove stroke (no line)
    shp.Line.Visible = msoFalse
End Sub
```

```javascript
// JavaScript code using OnlyOffice API to add a shape with a linear gradient fill and no stroke

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create gradient stops with RGB colors
var oGs1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0); // Light orange at position 0
var oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000); // Dark orange at position 100000

// Create a linear gradient fill with the gradient stops and angle
var oFill = Api.CreateLinearGradientFill([oGs1, oGs2], 5400000); // 5400000 represents the gradient angle

// Create a stroke with thickness 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add the shape to the worksheet
// Parameters: name, left, top, fill, stroke, rotation, width, height, other properties
oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000);
```