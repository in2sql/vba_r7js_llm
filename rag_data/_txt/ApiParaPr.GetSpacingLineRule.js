**Description / Описание**

*English:*  
This code creates a shape in the active worksheet with specified fill and stroke, adds a paragraph with specific line spacing, inserts text with line breaks, and retrieves the spacing line rule.

*Русский:*  
Этот код создает фигуру на активном листе с указанной заливкой и обводкой, добавляет абзац с определенным межстрочным расстоянием, вставляет текст с переносами строк и получает правило межстрочного интервала.

```javascript
// This example shows how to get the paragraph line spacing rule.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create no stroke
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000); // Add a shape to the worksheet
var oDocContent = oShape.GetContent(); // Get the shape's content
var oParagraph = oDocContent.GetElement(0); // Get the first paragraph
var oParaPr = oParagraph.GetParaPr(); // Get paragraph properties
oParaPr.SetSpacingLine(3 * 240, "auto"); // Set line spacing to 3 times the standard
oParagraph.AddText("Paragraph 1. Spacing: 3 times of a common paragraph line spacing."); // Add text to the paragraph
oParagraph.AddLineBreak(); // Add a line break
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes."); // Add more text
var sSpacingLineRule = oParaPr.GetSpacingLineRule(); // Retrieve the spacing line rule
oParagraph = Api.CreateParagraph(); // Create a new paragraph
oParagraph.AddText("Spacing line rule : " + sSpacingLineRule); // Add text displaying the spacing rule
oDocContent.Push(oParagraph); // Push the new paragraph to the document content
```

```vba
' This VBA code creates a shape in the active worksheet with specified fill and stroke,
' adds a paragraph with specific line spacing, inserts text with line breaks,
' and retrieves the spacing line rule.

Sub CreateShapeAndSetParagraphSpacing()
    ' Get the active worksheet
    Dim oSheet As Worksheet
    Set oSheet = ThisWorkbook.ActiveSheet

    ' Create a shape with specified properties
    Dim oShape As Shape
    Set oShape = oSheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, 120, 70, 200, 100) ' Position (120,70) and size (200x100) in points

    ' Set the fill color
    With oShape.Fill
        .ForeColor.RGB = RGB(255, 111, 61) ' RGB color equivalent to (255, 111, 61)
        .Solid
    End With

    ' Set no line (stroke)
    With oShape.Line
        .Visible = msoFalse
    End With

    ' Add text to the shape
    With oShape.TextFrame2
        .TextRange.Text = "Paragraph 1. Spacing: 3 times of a common paragraph line spacing." & vbCrLf & _
                          "These sentences are used to add lines for demonstrative purposes." ' Add text with a line break

        ' Set paragraph formatting
        With .TextRange.ParagraphFormat
            .LineSpacing = 3 * 12 ' Assuming standard line spacing is 12 points, set to 36 points
            .LineSpacingRule = msoLineSpacingMultiple ' Set line spacing rule to multiple
        End With
    End With

    ' Retrieve the spacing line rule
    Dim sSpacingLineRule As String
    With oShape.TextFrame2.TextRange.ParagraphFormat
        Select Case .LineSpacingRule
            Case msoLineSpacingSingle
                sSpacingLineRule = "Single"
            Case msoLineSpacing1pt5
                sSpacingLineRule = "1.5 lines"
            Case msoLineSpacingDouble
                sSpacingLineRule = "Double"
            Case msoLineSpacingAtLeast
                sSpacingLineRule = "At least"
            Case msoLineSpacingExactly
                sSpacingLineRule = "Exactly"
            Case msoLineSpacingMultiple
                sSpacingLineRule = "Multiple"
            Case Else
                sSpacingLineRule = "Unknown"
        End Select
    End With

    ' Add a new paragraph with the spacing line rule
    With oShape.TextFrame2.TextRange
        .Text = .Text & vbCrLf & "Spacing line rule : " & sSpacingLineRule
    End With
End Sub
```