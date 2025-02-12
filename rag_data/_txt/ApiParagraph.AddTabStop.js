## Description | Описание

**English:** This code adds a shape to the active sheet with specified fill and stroke, inserts a paragraph with text, adds three tab stops, and appends additional text after the tab stops.

**Russian:** Этот код добавляет фигуру на активный лист с указанной заливкой и обводкой, вставляет абзац с текстом, добавляет три табуляции и добавляет дополнительный текст после табуляций.

```vba
' VBA Code Equivalent

Sub AddShapeWithTabStops()
    Dim oWorksheet As Worksheet
    Dim oShape As Shape
    Dim oTextFrame As TextFrame
    Dim oTextRange As TextRange
    
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Add a shape to the worksheet
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartProcess, 120, 70, 200, 100)
    
    ' Set fill color using RGB
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61)
        .Solid
    End With
    
    ' Set no line (stroke)
    With oShape.Line
        .Visible = msoFalse
    End With
    
    ' Get the text frame of the shape
    Set oTextFrame = oShape.TextFrame
    
    ' Get the text range of the first paragraph
    Set oTextRange = oTextFrame.Characters
    
    ' Add initial text
    oTextRange.Text = "This is just a sample text. After it three tab stops will be added."
    
    ' Add three tab stops
    With oTextFrame
        .HorizontalAlignment = xlLeft
        .MarginLeft = 10
        .TextRange.Paragraphs.ParagraphFormat.TabStops.Add Position:=30
        .TextRange.Paragraphs.ParagraphFormat.TabStops.Add Position:=60
        .TextRange.Paragraphs.ParagraphFormat.TabStops.Add Position:=90
    End With
    
    ' Append text after tab stops
    oTextRange.Text = oTextRange.Text & vbTab & vbTab & vbTab & "This is the text which starts after the tab stops."
End Sub
```

```javascript
// OnlyOffice JavaScript Code Equivalent

// This example adds a shape to the active sheet with specified fill and stroke, 
// inserts a paragraph with text, adds three tab stops, and appends additional text after the tab stops.

var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Add initial text to the paragraph
oParagraph.AddText("This is just a sample text. After it three tab stops will be added.");

// Add three tab stops
oParagraph.AddTabStop();
oParagraph.AddTabStop();
oParagraph.AddTabStop();

// Add text after the tab stops
oParagraph.AddText("This is the text which starts after the tab stops.");
```