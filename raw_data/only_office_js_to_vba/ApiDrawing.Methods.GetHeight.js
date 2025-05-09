## Description / Описание

**English:**  
This code demonstrates how to add a flowchart shape to the active worksheet, set its fill color and stroke, adjust its size and position, retrieve its height, and display the height value in cell A1.

**Russian:**  
Этот код демонстрирует, как добавить форму блок-схемы на активный лист, установить цвет заливки и обводки, изменить размер и позицию, получить высоту формы и отобразить значение высоты в ячейке A1.

---

### VBA Code

```vba
' This VBA macro adds a flowchart shape to the active worksheet, sets its fill and stroke,
' adjusts its size and position, retrieves its height, and displays the height in cell A1.

Sub AddFlowchartShape()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim fillColor As Long
    Dim strokeColor As Long
    Dim strokeWeight As Single
    Dim shapeHeight As Single
    
    ' Get the active worksheet
    Set ws = ActiveSheet
    
    ' Define fill color using RGB
    fillColor = RGB(255, 111, 61)
    
    ' Define stroke properties
    strokeColor = RGB(0, 0, 0) ' Black color for stroke
    strokeWeight = 0 ' No stroke weight
    
    ' Add a flowchart shape (e.g., Flowchart: Database)
    Set shp = ws.Shapes.AddShape(Type:=msoShapeFlowchartDatabase, _
                                 Left:=60, Top:=35, Width:=120, Height:=70)
    
    ' Set the fill color
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = fillColor
        .Solid
    End With
    
    ' Set the stroke (line) properties
    With shp.Line
        .Visible = msoTrue
        .ForeColor.RGB = strokeColor
        .Weight = strokeWeight
    End With
    
    ' Set the position of the shape
    shp.Left = 0
    shp.Top = 2 * 36000 ' Adjusted as per requirement
    shp.LockAspectRatio = msoFalse
    
    ' Retrieve the height of the shape
    shapeHeight = shp.Height
    
    ' Display the height in cell A1
    ws.Range("A1").Value = "Drawing height = " & shapeHeight
End Sub
```

### OnlyOffice JS Code

```javascript
// This example shows how to add a flowchart shape to the active worksheet, set its fill and stroke,
// adjust its size and position, retrieve its height, and display the height in cell A1.

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with width 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a flowchart shape to the worksheet with specified position and size
var oDrawing = oWorksheet.AddShape(
    "flowChartOnlineStorage",
    60 * 36000,
    35 * 36000,
    oFill,
    oStroke,
    0,
    2 * 36000,
    0,
    3 * 36000
);

// Set the size of the shape
oDrawing.SetSize(120 * 36000, 70 * 36000);

// Set the position of the shape
oDrawing.SetPosition(0, 2 * 36000, 1, 3 * 36000);

// Get the height of the shape
var nHeight = oDrawing.GetHeight();

// Set the value of cell A1 to display the height
oWorksheet.GetRange("A1").SetValue("Drawing height = " + nHeight);
```