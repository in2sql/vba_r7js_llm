### Description / Описание

**English:**  
This code retrieves the active worksheet, creates a shape with a solid fill color and no stroke, clears any existing content within the shape, and adds a new left-aligned paragraph with specified text to the shape's content.

**Русский:**  
Этот код получает активный лист, создает фигуру со сплошной заливкой цвета и без обводки, очищает любое существующее содержимое внутри фигуры и добавляет новый абзац с установленным выравниванием влево и указанным текстом в содержимое фигуры.

```vba
' VBA Code

' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ActiveSheet

' Add a flowchart storage shape with specified position and size
Dim oShape As Shape
Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, 60, 35, 200, 100)

' Set the fill color to RGB(255, 111, 61)
With oShape.Fill
    .Solid
    .ForeColor.RGB = RGB(255, 111, 61)
End With

' Set the line (stroke) to no fill
oShape.Line.Visible = msoFalse

' Remove all existing text in the shape
oShape.TextFrame.Characters.Text = ""

' Add a new paragraph with left alignment and specified text
With oShape.TextFrame
    .HorizontalAlignment = xlHAlignLeft
    .Characters.Text = "We removed all elements from the shape and added a new paragraph inside it."
End With
```

```javascript
// OnlyOffice JS Code

// This example pushes 5 paragraphs to actually add its to the document content.
var oWorksheet = Api.GetActiveSheet(); // Get the active sheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill with RGB color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000); // Add a shape to the worksheet
var oDocContent = oShape.GetContent(); // Get the shape's content
oDocContent.RemoveAllElements(); // Remove all existing elements
var oParagraph = Api.CreateParagraph(); // Create a new paragraph
oParagraph.SetJc("left"); // Set alignment to left
oParagraph.AddText("We removed all elements from the shape and added a new paragraph inside it."); // Add text to paragraph
oDocContent.Push(oParagraph); // Add the paragraph to the content
```