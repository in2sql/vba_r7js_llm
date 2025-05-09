# Description / Описание

**English:**  
This code demonstrates how to create a shape in the active worksheet, add paragraphs with specific line spacing, and retrieve the line spacing value using the OnlyOffice API.

**Russian:**  
Этот код демонстрирует, как создать фигуру на активном листе, добавить абзацы с заданным межстрочным интервалом и получить значение межстрочного интервала с использованием OnlyOffice API.

```javascript
// This example shows how to get the paragraph line spacing value.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill with RGB color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with width 0 and no fill
// Add a shape to the worksheet with specified type, dimensions, fill, stroke, and positioning
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
var oDocContent = oShape.GetContent(); // Get the content of the shape
var oParagraph = oDocContent.GetElement(0); // Get the first paragraph
var oParaPr = oParagraph.GetParaPr(); // Get paragraph properties
oParaPr.SetSpacingLine(3 * 240, "auto"); // Set line spacing to 3 times the common spacing
// Add text to the paragraph
oParagraph.AddText("Paragraph 1. Spacing: 3 times of a common paragraph line spacing.");
oParagraph.AddLineBreak(); // Add a line break
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes."); // Add more text
var nSpacingLineValue = oParaPr.GetSpacingLineValue(); // Get the line spacing value
oParagraph = Api.CreateParagraph(); // Create a new paragraph
oParagraph.AddText("Spacing line value : " + nSpacingLineValue); // Add text with the spacing value
oDocContent.Push(oParagraph); // Push the new paragraph to the document content
```

```vba
' This VBA code creates a shape in the active worksheet, adds paragraphs with specific line spacing,
' and retrieves the line spacing value.

Sub CreateShapeWithParagraphs()
    ' Get the active worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Define RGB color components
    Dim red As Long, green As Long, blue As Long
    red = 255
    green = 111
    blue = 61
    
    ' Add a shape to the worksheet with specified type and dimensions
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartInternalStorage, 120, 70, 200, 100)
    
    ' Set the fill color of the shape
    shp.Fill.ForeColor.RGB = RGB(red, green, blue)
    
    ' Set the line (stroke) properties
    With shp.Line
        .Visible = msoFalse ' No stroke
    End With
    
    ' Access the text frame of the shape
    With shp.TextFrame2
        .TextRange.Text = "Paragraph 1. Spacing: 3 times of a common paragraph line spacing." & vbCrLf & _
                          "These sentences are used to add lines for demonstrative purposes."
                          
        ' Set paragraph formatting
        With .TextRange.ParagraphFormat
            .SpaceBefore = 0
            .SpaceAfter = 0
            .LineSpacing = 3 * 12 ' Assuming standard line spacing of 12 points
            .LineSpacingRule = msoLineSpacingExactly
        End With
        
        ' Retrieve the line spacing value
        Dim lineSpacingValue As Single
        lineSpacingValue = .TextRange.ParagraphFormat.LineSpacing
        
        ' Add a new paragraph with the spacing value
        .TextRange.Text = .TextRange.Text & vbCrLf & "Spacing line value: " & lineSpacingValue
    End With
End Sub
```