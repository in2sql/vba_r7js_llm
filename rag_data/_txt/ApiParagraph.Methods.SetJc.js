### **Description / Описание**
**English:** This code adds a shape to the active worksheet, inserts paragraphs with text aligned to center, right, and left.

**Русский:** Этот код добавляет фигуру на активный лист, вставляет абзацы с текстом, выровненным по центру, правому и левому краю.

```javascript
// This example sets the paragraph contents justification.
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified parameters
var oShape = oWorksheet.AddShape(
    "flowChartOnlineStorage",
    120 * 36000,
    70 * 36000,
    oFill,
    oStroke,
    0,
    2 * 36000,
    0,
    3 * 36000
);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Create and add a paragraph with center alignment
var oParagraph = oDocContent.GetElement(0);
oParagraph.AddText("This is a paragraph with the text in it aligned by the center. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.");
oParagraph.SetJc("center");

// Create a new paragraph with right alignment and add texts
oParagraph = Api.CreateParagraph();
oParagraph.AddText("This is a paragraph with the text in it aligned by the right side. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.");
oParagraph.SetJc("right");
oDocContent.Push(oParagraph);

// Create a new paragraph with left alignment and add texts
oParagraph = Api.CreateParagraph();
oParagraph.AddText("This is a paragraph with the text in it aligned by the left side. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.");
oParagraph.SetJc("left");
oDocContent.Push(oParagraph); 
```

```vba
' This VBA example sets the paragraph contents justification by adding a shape and aligning text
Sub AddShapeWithAlignedParagraphs()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Define RGB color
    Dim fillColor As Long
    fillColor = RGB(255, 111, 61)
    
    ' Add a shape to the worksheet
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowChartOnlineStorage, _
        120, 70, 200, 150)
    
    ' Set the fill color
    With oShape.Fill
        .Visible = msoTrue
        .Solid
        .ForeColor.RGB = fillColor
    End With
    
    ' Remove the line
    With oShape.Line
        .Visible = msoFalse
    End With
    
    ' Add and format paragraphs
    With oShape.TextFrame
        .HorizontalAlignment = xlCenter
        .Characters.Text = "This is a paragraph with the text in it aligned by the center. " & vbCrLf & _
                           "These sentences are used to add lines for demonstrative purposes. " & vbCrLf & _
                           "These sentences are used to add lines for demonstrative purposes."
    End With
    
    ' Add a second paragraph with right alignment
    With oShape.TextFrame2.TextRange
        .Paragraphs.Add
        .Paragraphs(2).ParagraphFormat.Alignment = msoAlignRight
        .Paragraphs(2).Text = "This is a paragraph with the text in it aligned by the right side. " & _
                              "These sentences are used to add lines for demonstrative purposes. " & _
                              "These sentences are used to add lines for demonstrative purposes."
    End With
    
    ' Add a third paragraph with left alignment
    With oShape.TextFrame2.TextRange
        .Paragraphs.Add
        .Paragraphs(3).ParagraphFormat.Alignment = msoAlignLeft
        .Paragraphs(3).Text = "This is a paragraph with the text in it aligned by the left side. " & _
                              "These sentences are used to add lines for demonstrative purposes. " & _
                              "These sentences are used to add lines for demonstrative purposes."
    End With
End Sub
```