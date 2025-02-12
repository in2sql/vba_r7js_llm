# Description / Описание

This code example demonstrates how to create a shape in the active worksheet, set its fill and stroke properties, and add text with specific line spacing. / Этот пример кода демонстрирует, как создать фигуру на активном листе, установить свойства заливки и обводки, а также добавить текст с определенным межстрочным интервалом.

```vba
' VBA Code

Sub AddShapeWithFormattedText()
    ' Get the active sheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Define fill color (RGB 255,111,61)
    Dim fillColor As Long
    fillColor = RGB(255, 111, 61)
    
    ' Add a shape - msoShapeFlowchartProcess is equivalent to "flowChartOnlineStorage"
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartProcess, 120, 80, 36000, 24000)
    
    ' Set fill color
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = fillColor
        .Solid
    End With
    
    ' Set no stroke
    With oShape.Line
        .Visible = msoFalse
    End With
    
    ' Add text to the shape
    With oShape.TextFrame2
        .WordWrap = msoTrue
        .AutoSize = msoAutoSizeNone
        With .TextRange
            ' Set line spacing to 3 times
            .ParagraphFormat.LineSpacingRule = msoLineSpacingMultiple
            .ParagraphFormat.LineSpacing = 3 * 240 ' Assuming 240 is the base unit
            
            ' Add text with line breaks
            .Text = "Paragraph 1. Spacing: 3 times of a common paragraph line spacing." & vbCrLf & _
                    "These sentences are used to add lines for demonstrative purposes. " & _
                    "These sentences are used to add lines for demonstrative purposes." & vbCrLf & _
                    "Spacing line value: 720" ' 3 * 240
        End With
    End With
End Sub
```

```javascript
// OnlyOffice JS Code

// This example shows how to create a shape, set fill and stroke, and add formatted text with line spacing.
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a flowchart storage shape with specified dimensions and styling
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 80 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Set line spacing to 3 times
oParagraph.SetSpacingLine(3 * 240, "auto");

// Add text to the paragraph
oParagraph.AddText("Paragraph 1. Spacing: 3 times of a common paragraph line spacing.");
oParagraph.AddLineBreak();
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddLineBreak();

// Get the line spacing value
var nSpacingLineValue = oParagraph.GetSpacingLineValue();

// Add the spacing line value to the text
oParagraph.AddText("Spacing line value: " + nSpacingLineValue); 
```