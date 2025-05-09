**Description / Описание**

This script creates an empty shape with specified fill and stroke, adds a paragraph with sample text, sets the font size to 30, and makes it bold.  
Этот скрипт создаёт пустую форму с заданной заливкой и обводкой, добавляет абзац с примерным текстом, устанавливает размер шрифта 30 и делает его жирным.

---

```javascript
// JavaScript Code using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified parameters
var oShape = oWorksheet.AddShape(
    "flowChartOnlineStorage", // Shape type
    80 * 36000,               // Left position
    50 * 36000,               // Top position
    oFill,                    // Fill
    oStroke,                  // Stroke
    0,                        // Rotation
    2 * 36000,                // Width
    0,                        // Height parameter not used here
    3 * 36000                 // Height
);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Remove all existing elements in the shape's content
oDocContent.RemoveAllElements();

// Create text properties
var oTextPr = Api.CreateTextPr();
oTextPr.SetFontSize(30); // Set font size to 30
oTextPr.SetBold(true);   // Set text to bold

// Create a new paragraph
var oParagraph = Api.CreateParagraph();
oParagraph.SetJc("left"); // Justify text to the left

// Add sample text to the paragraph
oParagraph.AddText("This is a sample text with the font size set to 30 and the font weight set to bold.");

// Apply text properties to the paragraph
oParagraph.SetTextPr(oTextPr);

// Add the paragraph to the shape's content
oDocContent.Push(oParagraph);
```

```vba
' VBA Code Equivalent

Sub CreateShapeWithText()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Define fill color RGB(255, 111, 61)
    Dim fillColor As Long
    fillColor = RGB(255, 111, 61)
    
    ' Add a shape to the worksheet
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape( _
        Type:=msoShapeFlowchartStorage, _ ' Shape type equivalent to "flowChartOnlineStorage"
        Left:=80 * 36000 / 36000, _       ' Left position in points
        Top:=50 * 36000 / 36000, _        ' Top position in points
        Width:=2 * 36000 / 36000, _       ' Width in points
        Height:=3 * 36000 / 36000)        ' Height in points
    
    ' Set the fill color
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = fillColor
        .Solid
    End With
    
    ' Remove any existing line
    With oShape.Line
        .Visible = msoFalse
    End With
    
    ' Add a text box inside the shape
    oShape.TextFrame2.TextRange.Text = ""
    
    ' Clear any existing text
    oShape.TextFrame2.TextRange.Text = ""
    
    ' Set text properties
    With oShape.TextFrame2.TextRange
        .Font.Size = 30     ' Set font size to 30
        .Font.Bold = msoTrue ' Set font to bold
        .ParagraphFormat.Alignment = msoAlignLeft ' Align text to left
        .Text = "This is a sample text with the font size set to 30 and the font weight set to bold."
    End With
End Sub
```