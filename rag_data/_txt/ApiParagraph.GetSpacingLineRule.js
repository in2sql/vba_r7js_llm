**Description / Описание:**

This script demonstrates how to add a shape with specific fill and stroke properties to the active worksheet, set the line spacing for paragraphs within the shape's content, add text and line breaks, and retrieve the current line spacing rule.

Этот скрипт демонстрирует, как добавить фигуру с определенными свойствами заливки и обводки на активный лист, установить межстрочный интервал для абзацев в содержимом фигуры, добавить текст и переносы строк, а также получить текущее правило межстрочного интервала.

```vba
' Excel VBA Code
Sub AddShapeAndSetParagraphSpacing()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Add a shape to the worksheet
    ' Using msoShapeFlowchartCalculation as an example since "flowChartOnlineStorage" is not a standard shape
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartCalculation, 120, 80, 120, 80)
    
    ' Set the fill color to RGB(255, 111, 61)
    oShape.Fill.ForeColor.RGB = RGB(255, 111, 61)
    oShape.Fill.Solid
    
    ' Set the line to no fill (no stroke)
    oShape.Line.Visible = msoFalse
    
    ' Access the text frame of the shape
    Dim oTextFrame As TextFrame2
    Set oTextFrame = oShape.TextFrame2
    
    ' Add text to the shape
    With oTextFrame.TextRange
        .Text = "Paragraph 1. Spacing: 3 times of a common paragraph line spacing." & vbCrLf & _
                "These sentences are used to add lines for demonstrative purposes. " & _
                "These sentences are used to add lines for demonstrative purposes. " & vbCrLf & _
                "Spacing line rule: "
        
        ' Set paragraph formatting for the first paragraph
        With .Paragraphs(1).ParagraphFormat
            .LineSpacingRule = msoLineSpacingExactly
            .LineSpacing = 3 * 12 ' Assuming 12 points per line
        End With
        
        ' Retrieve the line spacing rule
        Dim sSpacingLineRule As String
        Select Case .Paragraphs(1).ParagraphFormat.LineSpacingRule
            Case msoLineSpacingSingle
                sSpacingLineRule = "Single"
            Case msoLineSpacing1pt5
                sSpacingLineRule = "1.5 Lines"
            Case msoLineSpacingDouble
                sSpacingLineRule = "Double"
            Case msoLineSpacingAtLeast
                sSpacingLineRule = "At Least"
            Case msoLineSpacingExactly
                sSpacingLineRule = "Exactly"
            Case msoLineSpacingMultiple
                sSpacingLineRule = "Multiple"
            Case Else
                sSpacingLineRule = "Unknown"
        End Select
        
        ' Append the spacing line rule to the text
        .Text = .Text & sSpacingLineRule
    End With
End Sub
```

```javascript
// OnlyOffice JS Code
// This script adds a shape to the active worksheet with specific fill and stroke properties,
// sets paragraph line spacing, adds text and line breaks, and retrieves the current line spacing rule.

function AddShapeAndSetParagraphSpacing() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Create a solid fill with RGB color (255, 111, 61)
    var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
    
    // Create a stroke with 0 width and no fill
    var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    
    // Add a shape to the worksheet
    var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 80 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
    
    // Get the content of the shape
    var oDocContent = oShape.GetContent();
    
    // Get the first paragraph
    var oParagraph = oDocContent.GetElement(0);
    
    // Set the paragraph line spacing
    oParagraph.SetSpacingLine(3 * 240, "auto");
    
    // Add text to the paragraph
    oParagraph.AddText("Paragraph 1. Spacing: 3 times of a common paragraph line spacing.");
    oParagraph.AddLineBreak();
    oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
    oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
    oParagraph.AddLineBreak();
    
    // Get the spacing line rule
    var sSpacingLineRule = oParagraph.GetSpacingLineRule();
    
    // Add spacing line rule to the paragraph
    oParagraph.AddText("Spacing line rule: " + sSpacingLineRule);
}
```