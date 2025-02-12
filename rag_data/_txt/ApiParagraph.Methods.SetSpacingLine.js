### Description
**English:**  
This code sets the paragraph line spacing in an OnlyOffice worksheet by adding a shape and manipulating its content, including paragraphs with different line spacing and text.

**Russian:**  
Этот код устанавливает межстрочный интервал абзаца в листе OnlyOffice, добавляя фигуру и манипулируя её содержимым, включая абзацы с разным межстрочным интервалом и текстом.

### JavaScript Code
```javascript
// This example sets the paragraph line spacing.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill with RGB color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
// Add a shape to the worksheet with specified parameters
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
var oDocContent = oShape.GetContent(); // Get the content of the shape
var oParagraph = oDocContent.GetElement(0); // Get the first paragraph
oParagraph.SetSpacingLine(2 * 240, "auto"); // Set line spacing to double
oParagraph.AddText("Paragraph 1. Spacing: 2 times of a common paragraph line spacing."); // Add text to paragraph
oParagraph.AddLineBreak(); // Add a line break
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. "); // Add more text
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. "); // Add more text
oParagraph = Api.CreateParagraph(); // Create a new paragraph
oParagraph.SetSpacingLine(200, "exact"); // Set exact line spacing
oParagraph.AddText("Paragraph 2. Spacing: exact 10 points."); // Add text to new paragraph
oParagraph.AddLineBreak(); // Add a line break
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. "); // Add text
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. "); // Add text
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. "); // Add text
oDocContent.Push(oParagraph); // Add the new paragraph to the document content
```

### VBA Code
```vba
' This VBA code sets the paragraph line spacing by adding a shape and manipulating its text content.

Sub SetParagraphLineSpacing()
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet ' Get the active worksheet
    
    ' Create and add a shape to the worksheet
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowChartStorage, _
        120, 70, 200, 150) ' Position and size in points
    
    ' Set the fill color of the shape
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61) ' RGB color
        .Solid
    End With
    
    ' Remove the stroke (outline) of the shape
    With oShape.Line
        .Visible = msoFalse
    End With
    
    ' Access the text frame of the shape
    With oShape.TextFrame2.TextRange
        .Text = "Paragraph 1. Spacing: 2 times of a common paragraph line spacing." & vbCrLf & _
                "These sentences are used to add lines for demonstrative purposes. " & _
                "These sentences are used to add lines for demonstrative purposes. "
                
        ' Set line spacing to double
        .ParagraphFormat.LineSpacingRule = msoLineSpacingDouble
    End With
    
    ' Add a new paragraph with exact line spacing
    With oShape.TextFrame2.TextRange
        .InsertAfter vbCrLf & "Paragraph 2. Spacing: exact 10 points." & vbCrLf & _
            "These sentences are used to add lines for demonstrative purposes. " & _
            "These sentences are used to add lines for demonstrative purposes. " & _
            "These sentences are used to add lines for demonstrative purposes. "
        
        ' Set exact line spacing to 10 points
        Dim para As ParagraphFormat2
        Set para = .Paragraphs(.Paragraphs.Count).ParagraphFormat
        para.LineSpacingRule = msoLineSpacingExactly
        para.LineSpacing = 10 ' Points
    End With
End Sub
```