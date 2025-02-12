**Description / Описание**

This code adds a shape to the active sheet, sets its fill and stroke properties, and inserts paragraphs with specific line spacing and text.

Этот код добавляет фигуру на активный лист, устанавливает свойства заливки и обводки, а также вставляет абзацы с определённым межстрочным интервалом и текстом.

```vba
' VBA Code to add a shape with specific formatting and text

Sub AddFormattedShape()
    Dim oWorksheet As Worksheet
    Dim oShape As Shape
    Dim oFill As Variant
    Dim oStroke As Variant
    Dim oTextFrame As TextFrame2
    Dim oParagraph As TextRange2
    
    ' Set the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Define the fill color (RGB)
    Set oFill = oWorksheet.Shapes.AddShape(msoShapeRectangle, 100, 50, 300, 150).Fill
    oFill.ForeColor.RGB = RGB(255, 111, 61) ' RGB color
    
    ' Define the stroke (no fill)
    With oWorksheet.Shapes(1).Line
        .Visible = msoFalse
    End With
    
    ' Add a shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartStorage, 120, 70, 300, 150)
    oShape.Fill.ForeColor.RGB = RGB(255, 111, 61) ' Fill color
    oShape.Line.Visible = msoFalse ' No stroke
    
    ' Access the text frame of the shape
    Set oTextFrame = oShape.TextFrame2
    
    ' Add first paragraph with 2x line spacing
    Set oParagraph = oTextFrame.TextRange.Paragraphs.Add
    oParagraph.ParagraphFormat.SpaceWithin = 2
    oParagraph.Text = "Paragraph 1. Spacing: 2 times of a common paragraph line spacing." & vbCrLf
    oParagraph.Text = oParagraph.Text & "These sentences are used to add lines for demonstrative purposes. " & _
                      "These sentences are used to add lines for demonstrative purposes. "
    
    ' Add second paragraph with exact 10 points spacing
    Set oParagraph = oTextFrame.TextRange.Paragraphs.Add
    oParagraph.ParagraphFormat.SpaceAfter = 10 ' Points
    oParagraph.Text = "Paragraph 2. Spacing: exact 10 points." & vbCrLf
    oParagraph.Text = oParagraph.Text & "These sentences are used to add lines for demonstrative purposes. " & _
                      "These sentences are used to add lines for demonstrative purposes. " & _
                      "These sentences are used to add lines for demonstrative purposes."
End Sub
```

```javascript
// JavaScript Code to add a shape with specific formatting and text using OnlyOffice API

// This example sets the paragraph line spacing.
var oWorksheet = Api.GetActiveSheet(); // Get the active sheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create fill color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create no stroke
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000); // Add shape
var oDocContent = oShape.GetContent(); // Get content of the shape
var oParagraph = oDocContent.GetElement(0); // Get first paragraph
oParagraph.SetSpacingLine(2 * 240, "auto"); // Set line spacing to 2x
oParagraph.AddText("Paragraph 1. Spacing: 2 times of a common paragraph line spacing."); // Add text
oParagraph.AddLineBreak(); // Add line break
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. "); // Add text
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. "); // Add text
oParagraph = Api.CreateParagraph(); // Create new paragraph
oParagraph.SetSpacingLine(200, "exact"); // Set exact spacing of 10 points
oParagraph.AddText("Paragraph 2. Spacing: exact 10 points."); // Add text
oParagraph.AddLineBreak(); // Add line break
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. "); // Add text
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. "); // Add text
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. "); // Add text
oDocContent.Push(oParagraph); // Push the paragraph to content
```