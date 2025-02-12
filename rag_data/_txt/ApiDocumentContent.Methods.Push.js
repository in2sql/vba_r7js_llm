```plaintext
// Description: Creates a shape in the active worksheet, sets its fill and stroke, clears its content, and adds a new paragraph with text.
// Описание: Создает форму в активном листе, устанавливает ее заполнение и обводку, очищает содержимое и добавляет новый абзац с текстом.
```

```vba
' VBA Code
Sub AddShapeAndParagraph()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim fillColor As Long
    Dim leftPos As Single, topPos As Single, width As Single, height As Single

    ' Get the active worksheet
    Set ws = ActiveSheet

    ' Define fill color using RGB
    fillColor = RGB(255, 111, 61)

    ' Define position and size (scaled as needed)
    leftPos = 60 * 36    ' Excel uses points; adjust scaling if necessary
    topPos = 35 * 36
    width = 2 * 36
    height = 3 * 36

    ' Add a flowchart shape to the worksheet
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartOnlineStorage, leftPos, topPos, width, height)
    
    ' Set the fill color
    With shp.Fill
        .ForeColor.RGB = fillColor
        .Solid
    End With
    
    ' Remove any existing lines (stroke)
    With shp.Line
        .Visible = msoFalse
    End With
    
    ' Clear existing text
    shp.TextFrame2.TextRange.Text = ""
    
    ' Add new paragraph with text
    With shp.TextFrame2.TextRange
        .Text = "We removed all elements from the shape and added a new paragraph inside it."
        .ParagraphFormat.Alignment = msoAlignLeft
    End With
End Sub
```

```javascript
// OnlyOffice JS Code
// Creates a shape in the active worksheet, sets its fill and stroke, clears its content, and adds a new paragraph with text.
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a flowchart shape to the worksheet with specified dimensions
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the document content of the shape
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