**Description / Описание**

This code creates a stroke with shadows and adds a flowchart shape to the active worksheet using gradient and solid fills.
Этот код создает обводку с тенями и добавляет форму блок-схемы на активный лист с использованием градиентных и твердых заливок.

```vba
' VBA code to create a stroke with shadows and add a shape to the active worksheet

Sub AddShapeWithStroke()
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Define gradient colors
    Dim oColor1 As Long
    oColor1 = RGB(255, 213, 191)
    Dim oColor2 As Long
    oColor2 = RGB(255, 111, 61)
    
    ' Add a flowchart shape
    Dim shp As Shape
    Set shp = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, 60, 35, 100, 50)
    
    ' Apply gradient fill to the shape
    With shp.Fill
        .TwoColorGradient msoGradientHorizontal, 1
        .ForeColor.RGB = oColor1
        .BackColor.RGB = oColor2
    End With
    
    ' Set the stroke (line) properties
    With shp.Line
        .Weight = 3
        .ForeColor.RGB = RGB(51, 51, 51)
    End With
End Sub
```

```javascript
// This code creates a stroke adding shadows to the element and adds a flowchart shape to the active worksheet using OnlyOffice API

var oWorksheet = Api.GetActiveSheet();

// Create gradient stops
var oGs1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0);
var oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000);

// Create linear gradient fill
var oFill = Api.CreateLinearGradientFill([oGs1, oGs2], 5400000);

// Create solid fill for stroke
var oFill1 = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));

// Create stroke with width and fill
var oStroke = Api.CreateStroke(3 * 36000, oFill1);

// Add shape to worksheet
oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000);
```