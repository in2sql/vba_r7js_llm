### Description / Описание

**English:** This code adds a shape to the active worksheet, sets its fill and stroke, defines its size and position, retrieves its width, and writes the width value to cell A1.

**Russian:** Этот код добавляет фигуру на активный лист, устанавливает её заливку и обводку, задаёт её размер и позицию, получает её ширину и записывает значение ширины в ячейку A1.

```javascript
// This code adds a shape, sets its properties, retrieves its width, and writes it to cell A1

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill with specified RGB color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
var oDrawing = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000); // Add a shape to the worksheet
oDrawing.SetSize(120 * 36000, 70 * 36000); // Set the size of the shape
oDrawing.SetPosition(0, 2 * 36000, 1, 3 * 36000); // Set the position of the shape
var nWidth = oDrawing.GetWidth(); // Get the width of the shape
oWorksheet.GetRange("A1").SetValue("Drawing width = " + nWidth); // Write the width to cell A1
```

```vba
' This VBA code adds a shape, sets its properties, retrieves its width, and writes it to cell A1

Sub AddShapeAndGetWidth()
    Dim oWorksheet As Worksheet
    Dim oShape As Shape
    Dim nWidth As Double
    
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Add a shape to the worksheet
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, 60, 35, 120, 70)
    
    ' Set the fill color
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61) ' Set RGB color
        .Transparency = 0 ' No transparency
    End With
    
    ' Remove the stroke
    With oShape.Line
        .Visible = msoFalse ' No stroke
    End With
    
    ' Set the position of the shape
    oShape.Left = 0 ' Set left position
    oShape.Top = 2 ' Set top position (adjust units as needed)
    
    ' Get the width of the shape
    nWidth = oShape.Width
    
    ' Write the width to cell A1
    oWorksheet.Range("A1").Value = "Drawing width = " & nWidth
End Sub
```