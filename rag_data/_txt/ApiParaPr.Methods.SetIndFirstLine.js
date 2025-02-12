# Description / Описание

This code sets the first line indentation of a paragraph in a shape on an active worksheet, adds multiple lines of text, and styles the shape with specific fill and stroke colors.

Этот код устанавливает отступ первой строки абзаца в фигуре на активном листе, добавляет несколько строк текста и стилизует фигуру с использованием определенных цветов заливки и обводки.

```vba
' VBA Code to set paragraph first line indentation and add text to a shape

Sub SetParagraphIndentationAndAddText()
    Dim oWorksheet As Worksheet
    Dim oShape As Shape
    Dim oTextFrame As TextFrame
    Dim oParagraphFormat As ParagraphFormat
    
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Add a flowchart shape to the worksheet
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, _
        Left:=120 * 72, Top:=70 * 72, Width:=200, Height:=100)
    
    ' Set the fill color to RGB(255, 111, 61)
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61)
        .Solid
    End With
    
    ' Remove the stroke (no line)
    With oShape.Line
        .Visible = msoFalse
    End With
    
    ' Access the text frame of the shape
    Set oTextFrame = oShape.TextFrame
    
    ' Add text to the shape
    oTextFrame.Characters.Text = "This is the first paragraph with the indent of 1 inch set to the first line. " & _
        "This indent is set by the paragraph style. No paragraph inline style is applied. " & _
        "These sentences are used to add lines for demonstrative purposes. " & _
        "These sentences are used to add lines for demonstrative purposes. " & _
        "These sentences are used to add lines for demonstrative purposes."
    
    ' Set the first line indentation to 1 inch (72 points)
    With oTextFrame.Paragraphs(1).ParagraphFormat
        .FirstLineIndent = 72
    End With
End Sub
```

```javascript
// JavaScript code to set paragraph first line indentation and add text to a shape

// This example sets the paragraph first line indentation.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill with RGB color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with width 0 and no fill
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000); // Add a shape to the worksheet
var oDocContent = oShape.GetContent(); // Get the content of the shape
var oParagraph = oDocContent.GetElement(0); // Get the first paragraph
var oParaPr = oParagraph.GetParaPr(); // Get paragraph properties
oParaPr.SetIndFirstLine(1440); // Set first line indent to 1440 (points)
oParagraph.AddText("This is the first paragraph with the indent of 1 inch set to the first line. "); // Add text
oParagraph.AddText("This indent is set by the paragraph style. No paragraph inline style is applied. "); // Add text
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. "); // Add text
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. "); // Add text
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes."); // Add text
```