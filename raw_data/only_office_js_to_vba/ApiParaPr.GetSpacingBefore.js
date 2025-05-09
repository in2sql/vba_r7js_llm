### Description / Описание

**English:**  
This code demonstrates how to create a shape in the active worksheet, add text to paragraphs, set spacing before a paragraph, and retrieve the spacing value using OnlyOffice API and its Excel VBA equivalent.

**Russian:**  
Этот код демонстрирует, как создать фигуру на активном листе, добавить текст в абзацы, установить отступ перед абзацем и получить значение отступа, используя OnlyOffice API и эквивалентный код на Excel VBA.

```javascript
// JavaScript code using OnlyOffice API

// Get the active worksheet
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

// Add text to the first paragraph
oParagraph.AddText("This is an example of setting a space before a paragraph.");
oParagraph.AddText("The second paragraph will have an offset of one inch from the top. ");
oParagraph.AddText("This is due to the fact that the second paragraph has this offset enabled.");

// Create a second paragraph
var oParagraph2 = Api.CreateParagraph();
oParagraph2.AddText("This is the second paragraph and it is one inch away from the first paragraph.");

// Push the second paragraph to the document content
oDocContent.Push(oParagraph2);

// Get paragraph properties of the second paragraph
var oParaPr = oParagraph2.GetParaPr();

// Set spacing before the second paragraph
oParaPr.SetSpacingBefore(1440);

// Get the spacing before value
var nSpacingBefore = oParaPr.GetSpacingBefore();

// Create a new paragraph to display the spacing value
oParagraph = Api.CreateParagraph();
oParagraph.AddText("Spacing before: " + nSpacingBefore);

// Push the new paragraph to the document content
oDocContent.Push(oParagraph);
```

```vba
' VBA code equivalent using Excel object model

Sub CreateShapeAndSetSpacing()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Create a shape with specified parameters
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, _
        120, 70, 2 * 72, 3 * 72) ' Width and Height in points (1 inch = 72 points)
    
    ' Set the fill color to RGB(255, 111, 61)
    With oShape.Fill
        .ForeColor.RGB = RGB(255, 111, 61)
        .Solid
    End With
    
    ' Remove the line (stroke)
    oShape.Line.Visible = msoFalse
    
    ' Add text to the shape
    oShape.TextFrame.Characters.Text = "This is an example of setting a space before a paragraph." & vbCrLf & _
        "The second paragraph will have an offset of one inch from the top." & vbCrLf & _
        "This is due to the fact that the second paragraph has this offset enabled."
    
    ' Add a second paragraph with spacing before
    Dim oParagraph2 As TextRange
    Set oParagraph2 = oShape.TextFrame.Characters.Text & vbCrLf & _
        "This is the second paragraph and it is one inch away from the first paragraph."
    
    ' Note: Excel VBA does not support setting paragraph spacing directly.
    ' As a workaround, insert a line break with spaces or use multiple text boxes.
    ' Here, we'll append the spacing value as a separate text.
    oShape.TextFrame.Characters.Text = oShape.TextFrame.Characters.Text & vbCrLf & "Spacing before: 1440"
End Sub
```