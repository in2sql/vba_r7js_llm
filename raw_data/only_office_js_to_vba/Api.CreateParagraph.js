# Description / Описание

**English:** This code creates a new shape on the active sheet, removes all existing elements from the shape, and adds a new paragraph with specified text inside it.

**Russian:** Этот код создает новую фигуру на активном листе, удаляет все существующие элементы из фигуры и добавляет внутри неё новый абзац с указанным текстом.

# VBA Code

```vba
' This code creates a new shape on the active sheet, removes all existing elements from the shape, and adds a new paragraph with specified text inside it.
' Этот код создает новую фигуру на активном листе, удаляет все существующие элементы из фигуры и добавляет внутри неё новый абзац с указанным текстом.

Sub AddShapeAndParagraph()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim fillColor As Long
    Dim leftPos As Single, topPos As Single, widthVal As Single, heightVal As Single

    ' Get the active worksheet
    Set ws = ActiveSheet

    ' Define fill color (RGB 255, 111, 61)
    fillColor = RGB(255, 111, 61)

    ' Define position and size (adjust units as needed)
    leftPos = 60 ' Example value for Left position
    topPos = 35 ' Example value for Top position
    widthVal = 200 ' Example width
    heightVal = 100 ' Example height

    ' Add a flowchart storage shape to the worksheet
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartOnlineStorage, leftPos, topPos, widthVal, heightVal)

    ' Set the fill color of the shape
    shp.Fill.Solid
    shp.Fill.ForeColor.RGB = fillColor

    ' Remove all existing text in the shape
    shp.TextFrame.Characters.Text = ""

    ' Add new paragraph text with left alignment
    With shp.TextFrame2.TextRange
        .Text = "We removed all elements from the shape and added a new paragraph inside it."
        .ParagraphFormat.Alignment = msoAlignLeft
    End With
End Sub
```

# OnlyOffice JS Code

```javascript
// This code creates a new shape on the active sheet, removes all existing elements from the shape, and adds a new paragraph with specified text inside it.
// Этот код создает новую фигуру на активном листе, удаляет все существующие элементы из фигуры и добавляет внутри неё новый абзац с указанным текстом.

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with width 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape of type "flowChartOnlineStorage" with specified position and size
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Remove all existing elements from the shape's content
oDocContent.RemoveAllElements();

// Create a new paragraph
var oParagraph = Api.CreateParagraph();

// Set paragraph alignment to left
oParagraph.SetJc("left");

// Add text to the paragraph
oParagraph.AddText("We removed all elements from the shape and added a new paragraph inside it.");

// Push the new paragraph into the shape's content
oDocContent.Push(oParagraph);
```