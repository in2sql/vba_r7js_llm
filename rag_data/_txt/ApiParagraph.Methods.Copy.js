# Description / Описание

**English:** This code creates a shape on the active sheet with a specific fill and no stroke, adds text content with a line break, copies the paragraph, and appends the copied paragraph to the shape's text.

**Russian:** Этот код создает фигуру на активном листе с определенной заливкой и без обводки, добавляет текстовое содержимое с разрывом строки, копирует абзац и добавляет скопированный абзац в текст фигуры.

```vba
' VBA code to create a shape, add text, and copy a paragraph

Sub CreateShapeAndCopyParagraph()
    Dim oShape As Shape
    Dim oTextFrame As TextFrame
    Dim oTextRange As TextRange
    Dim oParagraph As TextRange
    Dim oParagraphCopy As TextRange
    
    ' Add a flow chart online storage shape to the active sheet
    Set oShape = ActiveSheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, 120, 70, 200, 100)
    
    ' Set the fill color to RGB(255, 111, 61)
    oShape.Fill.Solid
    oShape.Fill.ForeColor.RGB = RGB(255, 111, 61)
    
    ' Remove the line (no stroke)
    oShape.Line.Visible = msoFalse
    
    ' Get the text frame of the shape
    Set oTextFrame = oShape.TextFrame
    
    ' Set the text alignment to left
    oTextFrame.HorizontalAlignment = xlHAlignLeft
    
    ' Add text to the shape with a line break
    Set oTextRange = oTextFrame.Characters
    oTextRange.Text = "This is a text inside the shape aligned left." & vbCrLf & "This is a text after the line break."
    
    ' Get the first paragraph
    Set oParagraph = oTextRange.Paragraphs(1)
    
    ' Copy the paragraph
    Set oParagraphCopy = oParagraph.Duplicate
    
    ' Append the copied paragraph to the text frame
    oTextFrame.Characters.InsertAfter vbCrLf & oParagraphCopy.Text
End Sub
```

```javascript
// JavaScript code using OnlyOffice API to create a shape, add text, and copy a paragraph

// This example creates a paragraph copy.
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a "flowChartOnlineStorage" shape to the active sheet with specific position and size
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph in the content
var oParagraph = oDocContent.GetElement(0);

// Set the paragraph justification to left
oParagraph.SetJc("left");

// Add text to the paragraph
oParagraph.AddText("This is a text inside the shape aligned left.");

// Add a line break
oParagraph.AddLineBreak();

// Add more text after the line break
oParagraph.AddText("This is a text after the line break.");

// Copy the paragraph
var oParagraph2 = oParagraph.Copy();

// Push the copied paragraph to the content
oDocContent.Push(oParagraph2);
```