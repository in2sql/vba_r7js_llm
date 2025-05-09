### Description / Описание
This code adds a shape to the active worksheet, sets its fill and stroke, adds a paragraph of text, removes all existing elements, and then adds a new paragraph with updated text.
Этот код добавляет фигуру на активный рабочий лист, устанавливает её заливку и обводку, добавляет абзац текста, удаляет все существующие элементы, а затем добавляет новый абзац с обновленным текстом.

```vba
' VBA Code to manipulate shapes and text in Excel

Sub ModifyShapeAndText()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Create a shape with specified dimensions
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, 60, 35, 200, 150)
    
    ' Set the fill color of the shape
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61)
        .Solid
    End With
    
    ' Set the line (stroke) properties of the shape
    With oShape.Line
        .Visible = msoTrue
        .Weight = 0
        .ForeColor.RGB = RGB(255, 255, 255) ' No fill equivalent
    End With
    
    ' Add text to the shape
    With oShape.TextFrame2.TextRange
        .Text = "This is just a sample paragraph."
        ' Remove all existing text
        .Delete
        ' Add new paragraph with left alignment
        .Text = "We removed all elements from the shape and added a new paragraph inside it."
        .ParagraphFormat.Alignment = msoAlignLeft
    End With
End Sub
```

```javascript
// JavaScript Code to manipulate shapes and text using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified parameters
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Add text to the paragraph
oParagraph.AddText("This is just a sample paragraph.");

// Remove all elements from the document content
oDocContent.RemoveAllElements();

// Create a new paragraph
oParagraph = Api.CreateParagraph();

// Set justification to left
oParagraph.SetJc("left");

// Add new text to the paragraph
oParagraph.AddText("We removed all elements from the shape and added a new paragraph inside it.");

// Push the new paragraph into the document content
oDocContent.Push(oParagraph);
```