## Description / Описание

**English:**  
This code creates a linear gradient fill with two gradient stops and adds a flowchart shape to the active worksheet with the specified fill and no stroke.

**Русский:**  
Этот код создает линейную градиентную заливку с двумя градиентными точками и добавляет форму блок-схемы на активный лист с указанной заливкой и без обводки.

```javascript
// English: Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// English: Create the first gradient stop with RGB color (255, 213, 191) at position 0
var oGs1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0);

// English: Create the second gradient stop with RGB color (255, 111, 61) at position 100000
var oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000);

// English: Create a linear gradient fill with the two gradient stops and an angle of 5400000
var oFill = Api.CreateLinearGradientFill([oGs1, oGs2], 5400000);

// English: Create a stroke with width 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// English: Add a shape named "flowChartOnlineStorage" to the worksheet with specified dimensions and styles
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

```vba
' English: Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ThisWorkbook.ActiveSheet

' English: Add a flowchart shape to the worksheet
Dim shp As Shape
Set shp = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, _
    60 * 36000, _ ' Left position
    35 * 36000, _ ' Top position
    2 * 36000, _  ' Width
    3 * 36000)    ' Height

' English: Define the first RGB color for the gradient (255, 213, 191)
Dim color1 As Long
color1 = RGB(255, 213, 191)

' English: Define the second RGB color for the gradient (255, 111, 61)
Dim color2 As Long
color2 = RGB(255, 111, 61)

' English: Set the fill of the shape to a linear gradient
With shp.Fill
    .Visible = msoTrue
    .ForeColor.RGB = color1
    .BackColor.RGB = color2
    .TwoColorGradient msoGradientHorizontal, 1
    .GradientStops(1).Position = 0
    .GradientStops(2).Position = 1
    .GradientAngle = 5400000 ' Note: VBA gradient angle is limited to 0-360
End With

' English: Remove the outline (stroke) of the shape
With shp.Line
    .Visible = msoFalse
End With
```