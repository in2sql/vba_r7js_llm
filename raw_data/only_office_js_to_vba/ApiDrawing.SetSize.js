**Description / Описание**

*English:* This script adds a "flowChartOnlineStorage" shape to the active worksheet, sets its fill color and stroke, and defines its size and position.

*Russian:* Этот скрипт добавляет форму "flowChartOnlineStorage" на активный лист, устанавливает цвет заливки и обводки, а также определяет ее размер и позицию.

```vba
' VBA code to add and configure a shape on the active worksheet

Sub AddFlowChartOnlineStorageShape()
    Dim oWorksheet As Worksheet
    Dim oShape As Shape
    Dim fillColor As Long
    
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Define fill color RGB(255, 111, 61)
    fillColor = RGB(255, 111, 61)
    
    ' Add the "flowChartOnlineStorage" shape with specified position and size
    Set oShape = oWorksheet.Shapes.AddShape( _
        Type:=msoShapeFlowchartAlternateProcess, _ ' Replace with appropriate shape type
        Left:=60 * 36000, _ 
        Top:=35 * 36000, _ 
        Width:=120 * 36000, _ 
        Height:=70 * 36000)
    
    ' Set fill color
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = fillColor
        .Solid
    End With
    
    ' Set stroke to no fill
    With oShape.Line
        .Visible = msoFalse
    End With
    
    ' Set position
    oShape.Left = 0
    oShape.Top = 2 * 36000
End Sub
```

```javascript
// JavaScript code to add and configure a shape using OnlyOffice API

// This example sets the size of the shape bounding box.
var oWorksheet = Api.GetActiveSheet();
// Create a solid fill color with RGB(255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
// Create a stroke with width 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
// Add the "flowChartOnlineStorage" shape with specified parameters
var oDrawing = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
// Set the size of the shape
oDrawing.SetSize(120 * 36000, 70 * 36000);
// Set the position of the shape
oDrawing.SetPosition(0, 2 * 36000, 2, 3 * 36000);
```