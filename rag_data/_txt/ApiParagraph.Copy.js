# Description / Описание

This code creates a paragraph copy inside a shape in the active worksheet, sets its alignment, adds text and line breaks, and pushes the copied paragraph into the shape's content.
Этот код создает копию абзаца внутри фигуры на активном листе, устанавливает его выравнивание, добавляет текст и разрывы строк, и добавляет скопированный абзац в содержимое фигуры.

```vba
' VBA code to create a paragraph copy inside a shape in the active worksheet

Sub CreateParagraphCopy()
    Dim oWorksheet As Worksheet
    Dim oShape As Shape
    Dim oParagraph As TextRange
    Dim oParagraph2 As TextRange

    ' Get the active worksheet
    Set oWorksheet = ActiveSheet

    ' Add a shape to the worksheet
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartStorage, _
        120, 70, 200, 100) ' Parameters: Type, Left, Top, Width, Height

    ' Access the text frame of the shape
    With oShape.TextFrame2
        ' Set horizontal alignment to left
        .TextRange.ParagraphFormat.Alignment = msoAlignLeft
        
        ' Add first line of text
        .TextRange.Text = "This is a text inside the shape aligned left."
        
        ' Add a line break and second line of text
        .TextRange.Text = .TextRange.Text & vbCrLf & "This is a text after the line break."
        
        ' Copy the paragraph
        Set oParagraph = .TextRange
        Set oParagraph2 = oParagraph.Duplicate
        
        ' Add the copied paragraph to the text frame
        .TextRange.Text = .TextRange.Text & vbCrLf & oParagraph2.Text
    End With
End Sub
```

```javascript
// JavaScript code to create a paragraph copy inside a shape in the active worksheet

// This example creates a paragraph copy.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill with RGB color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000); // Add a shape to the worksheet
var oDocContent = oShape.GetContent(); // Get the content of the shape
var oParagraph = oDocContent.GetElement(0); // Get the first paragraph
oParagraph.SetJc("left"); // Set alignment to left
oParagraph.AddText("This is a text inside the shape aligned left."); // Add text
oParagraph.AddLineBreak(); // Add a line break
oParagraph.AddText("This is a text after the line break."); // Add more text
var oParagraph2 = oParagraph.Copy(); // Copy the paragraph
oDocContent.Push(oParagraph2); // Push the copied paragraph into the content
```