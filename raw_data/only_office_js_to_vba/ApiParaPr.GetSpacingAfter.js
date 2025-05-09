**Description / Описание**

English:  
This code adds a shape to the active worksheet, sets its fill and stroke properties, modifies paragraph spacing, adds multiple lines of text to the paragraph, and displays the spacing after the paragraph.

Русский:  
Этот код добавляет фигуру на активный лист, устанавливает свойства заливки и обводки, изменяет интервал между абзацами, добавляет несколько строк текста в абзац и отображает интервал после абзаца.

```vba
' VBA Code
Sub AddShapeAndModifyParagraph()
    Dim oWorksheet As Worksheet
    Dim oShape As Shape
    Dim oTextFrame As TextFrame2
    Dim oParagraph As ParagraphFormat
    Dim nSpacingAfter As Single

    ' Get the active worksheet
    Set oWorksheet = ActiveSheet

    ' Add a shape to the worksheet
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, _
        120, 70, 200, 100) ' Width and Height are in points

    ' Set the fill color (RGB: 255, 111, 61)
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61)
        .Solid
    End With

    ' Set the line (stroke) to no fill
    With oShape.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0) ' Default line color
        .Weight = 0 ' No line
    End With

    ' Access the text frame of the shape
    Set oTextFrame = oShape.TextFrame2

    ' Access the first paragraph
    With oTextFrame.TextRange.ParagraphFormat
        ' Set spacing after to 1440 (points)
        .SpaceAfter = 1440
    End With

    ' Add multiple lines of text
    With oTextFrame.TextRange
        .Text = "This is an example of setting a space after a paragraph." & vbCrLf
        .Text = .Text & "The second paragraph will have an offset of one inch from the top." & vbCrLf
        .Text = .Text & "These sentences are used to add lines for demonstrative purposes." & vbCrLf
        .Text = .Text & "These sentences are used to add lines for demonstrative purposes." & vbCrLf
        .Text = .Text & "These sentences are used to add lines for demonstrative purposes."
    End With

    ' Get the spacing after value
    nSpacingAfter = oTextFrame.TextRange.ParagraphFormat.SpaceAfter

    ' Add a new paragraph to display spacing after
    With oTextFrame.TextRange
        .Paragraphs.Add
        .Paragraphs(oTextFrame.TextRange.Paragraphs.Count).Text = "Spacing after: " & nSpacingAfter
    End With
End Sub
```

```javascript
// OnlyOffice JS Code
// This example shows how to get the spacing after value of the current paragraph.
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill color with RGB (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified dimensions and styles
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Get the paragraph properties
var oParaPr = oParagraph.GetParaPr();

// Set spacing after to 1440
oParaPr.SetSpacingAfter(1440);

// Add multiple lines of text to the paragraph
oParagraph.AddText("This is an example of setting a space after a paragraph. ");
oParagraph.AddText("The second paragraph will have an offset of one inch from the top. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.");

// Get the spacing after value
var nSpacingAfter = oParaPr.GetSpacingAfter();

// Create a new paragraph
oParagraph = Api.CreateParagraph();

// Add text displaying the spacing after value
oParagraph.AddText("Spacing after : " + nSpacingAfter);

// Push the new paragraph to the document content
oDocContent.Push(oParagraph);
```