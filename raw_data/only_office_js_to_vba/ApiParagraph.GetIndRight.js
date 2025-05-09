```plaintext
// This code adds a shape to the active sheet, sets fill and stroke, adds a paragraph with right indent of 2 inches (2880 twips), and displays the indent value.
// Этот код добавляет фигуру на активный лист, устанавливает заливку и обводку, добавляет параграф с отступом справа 2 дюйма (2880 twips) и отображает значение отступа.
```

```vba
' VBA code to add a shape with right indentation and display the indent value
Sub AddShapeWithIndent()
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Create a shape with specified type and dimensions
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, 120, 70, 120, 70)
    
    ' Set fill color to RGB(255, 111, 61)
    oShape.Fill.ForeColor.RGB = RGB(255, 111, 61)
    
    ' Remove the stroke
    oShape.Line.Visible = msoFalse
    
    ' Add text to the shape
    With oShape.TextFrame2.TextRange
        .Text = "This is a paragraph with the right offset of 2 inches set to it. " & _
                "These sentences are used to add lines for demonstrative purposes. "
        ' Set paragraph alignment to right
        .ParagraphFormat.Alignment = msoAlignRight
        ' Set right indentation to 144 points (2 inches)
        .ParagraphFormat.RightIndent = 144
        ' Retrieve the right indent value
        Dim nIndRight As Single
        nIndRight = .ParagraphFormat.RightIndent
        ' Add a new paragraph with the indent value
        .InsertAfter vbCrLf & "Right indent: " & nIndRight
    End With
End Sub
```

```javascript
// This code adds a shape to the active sheet, sets fill and stroke, adds a paragraph with right indent of 2 inches (2880 twips), and displays the indent value.
// Этот код добавляет фигуру на активный лист, устанавливает заливку и обводку, добавляет параграф с отступом справа 2 дюйма (2880 twips) и отображает значение отступа.

var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with width 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified parameters
var oShape = oWorksheet.AddShape(
    "flowChartOnlineStorage",
    120 * 36000, // Left position
    70 * 36000,  // Top position
    oFill,       // Fill
    oStroke,     // Stroke
    0,           // Width
    2 * 36000,   // Height
    0,           // Rotation
    3 * 36000    // Other parameters as needed
);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Add text to the paragraph
oParagraph.AddText("This is a paragraph with the right offset of 2 inches set to it. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");

// Set paragraph alignment to right
oParagraph.SetJc("right");

// Set right indentation to 2880 twips (2 inches)
oParagraph.SetIndRight(2880);

// Retrieve the right indentation value
var nIndRight = oParagraph.GetIndRight();

// Create a new paragraph to display the indentation value
oParagraph = Api.CreateParagraph();
oParagraph.AddText("Right indent: " + nIndRight);

// Add the new paragraph to the document content
oDocContent.Push(oParagraph);
```