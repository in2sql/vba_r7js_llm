# Description / Описание

**English:** This code adds a shape to the active worksheet, sets its size and position, retrieves its height, and writes the height value to cell A1.

**Russian:** Этот код добавляет фигуру на активный лист, устанавливает ее размер и позицию, получает ее высоту и записывает значение высоты в ячейку A1.

```vba
' VBA Code
' This code adds a shape to the active worksheet, sets its size and position,
' retrieves its height, and writes the height value to cell A1.

Sub AddShapeAndGetHeight()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim nHeight As Single
    
    ' Set active worksheet
    Set ws = ActiveSheet
    
    ' Add a flowchart shape with specified parameters
    Set shp = ws.Shapes.AddShape(Type:=msoShapeFlowchartOnlineStorage, _
                                 Left:=60 * 36000, Top:=35 * 36000, Width:=120 * 36000, Height:=70 * 36000)
    
    ' Set position
    shp.Left = 0
    shp.Top = 2 * 36000
    
    ' Get height of the shape
    nHeight = shp.Height
    
    ' Write height to cell A1
    ws.Range("A1").Value = "Drawing height = " & nHeight
End Sub
```

```javascript
// OnlyOffice JS Code
// This code adds a shape to the active worksheet, sets its size and position,
// retrieves its height, and writes the height value to cell A1.

var oWorksheet = Api.GetActiveSheet();
// Create solid fill with RGB color
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
// Create stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
// Add a flowchart shape with specified parameters
var oDrawing = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
// Set size of the drawing
oDrawing.SetSize(120 * 36000, 70 * 36000);
// Set position of the drawing
oDrawing.SetPosition(0, 2 * 36000, 1, 3 * 36000);
// Get height of the drawing
var nHeight = oDrawing.GetHeight();
// Write height to cell A1
oWorksheet.GetRange("A1").SetValue("Drawing height = " + nHeight); 
```