# Description / Описание

**English:**  
This code adds a shape to the active worksheet, sets its fill and stroke properties, and adds a text run to the first paragraph of the shape's content.

**Russian:**  
Этот код добавляет фигуру на активный лист, устанавливает свойства заливки и обводки, а затем добавляет текстовый элемент в первый параграф содержимого фигуры.

```javascript
// JavaScript (OnlyOffice) Code
// This example adds a Run to the paragraph.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill with RGB color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000); // Add shape to worksheet
var oDocContent = oShape.GetContent(); // Get the document content of the shape
var oParagraph = oDocContent.GetElement(0); // Get the first paragraph
var oRun = Api.CreateRun(); // Create a text run
oRun.AddText("This is just a sample text run. Nothing special."); // Add text to the run
oParagraph.AddElement(oRun); // Add the run to the paragraph
```

```vba
' Excel VBA Code
' This example adds a shape to the active worksheet and adds a text run to the first paragraph.

Sub AddShapeAndRun()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim fillColor As Long
    
    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Define fill color (RGB)
    fillColor = RGB(255, 111, 61)
    
    ' Add a shape to the worksheet (example uses a flowchart database shape)
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartDatabase, 120, 70, 200, 100) ' Adjust dimensions and position as needed
    
    ' Set the fill color
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = fillColor
        .Solid
    End With
    
    ' Set the stroke to have no fill (VBA does not support 'no fill' directly, so setting weight to 0)
    With shp.Line
        .Visible = msoFalse
    End With
    
    ' Add text to the shape
    shp.TextFrame2.TextRange.Text = "This is just a sample text run. Nothing special."
    
End Sub
```