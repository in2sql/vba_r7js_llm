**Description / Описание**

This code adds a shape to the active worksheet, sets its fill and stroke, adds paragraphs with specific line spacing, and retrieves the spacing line rule.
Этот код добавляет фигуру на активный лист, устанавливает ее заливку и обводку, добавляет абзацы с определенным межстрочным интервалом и извлекает правило межстрочного интервала.

```vba
' VBA Code to add a shape, set fill and stroke, add paragraphs with specific line spacing, and retrieve spacing line rule

Sub AddShapeAndParagraphs()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Define fill color (RGB: 255, 111, 61)
    Dim fillColor As Long
    fillColor = RGB(255, 111, 61)
    
    ' Add a shape to the worksheet
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartAlternateProcess, 120, 70, 120, 70)
    
    ' Set the fill color
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = fillColor
        .Solid
    End With
    
    ' Set the stroke (no fill)
    With oShape.Line
        .Visible = msoFalse
    End With
    
    ' Add text to the shape
    Dim sText As String
    sText = "Paragraph 1. Spacing: 3 times of a common paragraph line spacing." & vbCrLf & _
            "These sentences are used to add lines for demonstrative purposes."
    oShape.TextFrame2.TextRange.Text = sText
    
    ' Set paragraph line spacing to 3 times the common spacing
    With oShape.TextFrame2.TextRange.ParagraphFormat
        .LineSpacingRule = msoLineSpacingMultiple
        .LineSpacing = 3 * 12 ' Assuming 12 points as the common spacing
    End With
    
    ' Retrieve the line spacing rule
    Dim spacingRule As String
    Select Case oShape.TextFrame2.TextRange.ParagraphFormat.LineSpacingRule
        Case msoLineSpacingSingle
            spacingRule = "Single"
        Case msoLineSpacing1pt5
            spacingRule = "1.5 Lines"
        Case msoLineSpacingDouble
            spacingRule = "Double"
        Case msoLineSpacingAtLeast
            spacingRule = "At Least"
        Case msoLineSpacingExactly
            spacingRule = "Exactly"
        Case msoLineSpacingMultiple
            spacingRule = "Multiple"
        Case Else
            spacingRule = "Unknown"
    End Select
    
    ' Add a new paragraph to display the spacing rule
    Dim oNewShape As Shape
    Set oNewShape = oWorksheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 300, 70, 300, 50)
    oNewShape.TextFrame2.TextRange.Text = "Spacing line rule : " & spacingRule
End Sub
```

```javascript
// JavaScript Code to add a shape, set fill and stroke, add paragraphs with specific line spacing, and retrieve spacing line rule

// Get the active sheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified parameters
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Get paragraph properties
var oParaPr = oParagraph.GetParaPr();

// Set line spacing to 3 times the common spacing with "auto" rule
oParaPr.SetSpacingLine(3 * 240, "auto");

// Add text to the paragraph
oParagraph.AddText("Paragraph 1. Spacing: 3 times of a common paragraph line spacing.");
oParagraph.AddLineBreak();
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.");

// Retrieve the spacing line rule
var sSpacingLineRule = oParaPr.GetSpacingLineRule();

// Create a new paragraph to display the spacing rule
oParagraph = Api.CreateParagraph();
oParagraph.AddText("Spacing line rule : " + sSpacingLineRule);

// Push the new paragraph to the document content
oDocContent.Push(oParagraph);
```