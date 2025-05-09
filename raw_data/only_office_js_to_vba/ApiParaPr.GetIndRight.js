```plaintext
// Description: This script demonstrates how to create a shape with specific fill and stroke, set paragraph indentation and alignment, and add text to the shape's content.
// Описание: Этот скрипт демонстрирует, как создать фигуру с определенной заливкой и обводкой, установить отступ и выравнивание абзаца и добавить текст в содержание фигуры.
```

```vba
' VBA Code: Create a shape, set paragraph indentation and alignment, and add text to the shape's content

Sub CreateShapeWithParagraph()
    Dim oWorksheet As Worksheet
    Dim oShape As Shape
    Dim oTextFrame As TextFrame
    Dim oParagraph As TextRange
    Dim nIndRight As Single
    
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Add a shape to the worksheet
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, _
                                           120, 70, 2, 3) ' Width and Height in points
    
    ' Set the fill color (RGB: 255, 111, 61)
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61)
        .Solid
    End With
    
    ' Set the line (stroke) to no fill
    With oShape.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 255) ' No fill equivalent
        .Weight = 0
    End With
    
    ' Access the text frame of the shape
    Set oTextFrame = oShape.TextFrame
    
    ' Clear any existing text
    oTextFrame.TextRange.Text = ""
    
    ' Add first paragraph with right indentation and alignment
    Set oParagraph = oTextFrame.TextRange.Paragraphs.Add
    With oParagraph
        .ParagraphFormat.RightIndent = Application.CentimetersToPoints(5.08) ' 2 inches
        .ParagraphFormat.Alignment = xlRight
        .Text = "This is the first paragraph with the right offset of 2 inches set to it. " & _
                "This indent is set by the paragraph style. No paragraph inline style is applied."
    End With
    
    ' Get the right indentation value
    nIndRight = oParagraph.ParagraphFormat.RightIndent
    
    ' Add second paragraph displaying the right indent value
    Set oParagraph = oTextFrame.TextRange.Paragraphs.Add
    oParagraph.Text = "Right indent: " & Format(nIndRight, "0.00") & " points"
End Sub
```

```javascript
// JavaScript Code: Create a shape, set paragraph indentation and alignment, and add text to the shape's content.

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with width 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a flow chart shape to the worksheet with specified dimensions and styles
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape's document
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Get paragraph properties
var oParaPr = oParagraph.GetParaPr();

// Set right indentation to 2880 (points)
oParaPr.SetIndRight(2880);

// Set paragraph alignment to right
oParaPr.SetJc("right");

// Add text to the paragraph
oParagraph.AddText("This is the first paragraph with the right offset of 2 inches set to it. ");
oParagraph.AddText("This indent is set by the paragraph style. No paragraph inline style is applied. ");

// Get the right indentation value
var nIndRight = oParaPr.GetIndRight();

// Create a new paragraph
oParagraph = Api.CreateParagraph();

// Add text displaying the right indentation value
oParagraph.AddText("Right indent: " + nIndRight);

// Push the new paragraph to the document content
oDocContent.Push(oParagraph);
```