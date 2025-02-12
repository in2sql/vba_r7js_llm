**Description / Описание:**

This code adds a flow chart shape to the active sheet, removes all its existing elements, adds a new paragraph with specific text, and sets its fill and stroke properties.

Этот код добавляет форму блок-схемы на активный лист, удаляет все ее существующие элементы, добавляет новый абзац с определенным текстом и устанавливает свойства заливки и обводки.

```vba
' VBA Code
Sub AddFlowChartShape()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' Get the active sheet
    
    ' Define fill color (RGB: 255, 111, 61)
    Dim fillColor As Long
    fillColor = RGB(255, 111, 61)
    
    ' Add a flowchart shape to the sheet
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartOnlineStorage, 200, 60, 200, 60)
    
    ' Set the fill color
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = fillColor
        .Solid
    End With
    
    ' Set the line (stroke) properties
    With shp.Line
        .Weight = 0 ' No stroke weight
        .Visible = msoFalse ' No stroke
    End With
    
    ' Remove existing text
    shp.TextFrame.Characters.Text = ""
    
    ' Add new text
    shp.TextFrame.Characters.Text = "We removed all elements from the shape and added a new paragraph inside it."
End Sub
```

```javascript
// OnlyOffice JS Code
// This code adds a flow chart shape to the active sheet, removes all its existing elements, adds a new paragraph with specific text, and sets its fill and stroke properties.
var oWorksheet = Api.GetActiveSheet(); // Get the active sheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create solid fill with RGB color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create stroke with weight 0 and no fill
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 200 * 36000, 60 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000); // Add shape to worksheet
var oDocContent = oShape.GetContent(); // Get shape content
oDocContent.RemoveAllElements(); // Remove all existing elements
var oParagraph = Api.CreateParagraph(); // Create new paragraph
oParagraph.AddText("We removed all elements from the shape and added a new paragraph inside it."); // Add text to paragraph
oDocContent.AddElement(oParagraph); // Add paragraph to content
oDocContent.Push(oParagraph); // Push paragraph to content
```