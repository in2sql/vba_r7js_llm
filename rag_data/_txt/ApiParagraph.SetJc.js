```plaintext
/*
Description: This code replicates the OnlyOffice JS example by setting paragraph alignments in an Excel worksheet.
Описание: Этот код воспроизводит пример OnlyOffice JS, устанавливая выравнивание абзацев в рабочем листе Excel.
*/
```

```vba
' VBA Code
Sub SetParagraphAlignments()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Add a shape to the worksheet
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartPredefined, 120, 70, 200, 100)
    
    ' Add text to the shape
    With oShape.TextFrame2
        ' Center aligned paragraph
        .TextRange.Text = "This is a paragraph with the text in it aligned by the center. " & _
                          "These sentences are used to add lines for demonstrative purposes. " & _
                          "These sentences are used to add lines for demonstrative purposes."
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        
        ' Add a new paragraph with right alignment
        .TextRange.Text = .TextRange.Text & vbCrLf & _
                          "This is a paragraph with the text in it aligned by the right side. " & _
                          "These sentences are used to add lines for demonstrative purposes. " & _
                          "These sentences are used to add lines for demonstrative purposes."
        .TextRange.Paragraphs(2).ParagraphFormat.Alignment = msoAlignRight
        
        ' Add another paragraph with left alignment
        .TextRange.Text = .TextRange.Text & vbCrLf & _
                          "This is a paragraph with the text in it aligned by the left side. " & _
                          "These sentences are used to add lines for demonstrative purposes. " & _
                          "These sentences are used to add lines for demonstrative purposes."
        .TextRange.Paragraphs(3).ParagraphFormat.Alignment = msoAlignLeft
    End With
    
    ' Set fill color for the shape
    oShape.Fill.ForeColor.RGB = RGB(255, 111, 61)
    
    ' Remove the shape's line
    oShape.Line.Visible = msoFalse
End Sub
```

```javascript
// OnlyOffice JS Code
// This example sets the paragraph contents justification.

var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph
var oParagraph = oDocContent.GetElement(0);

// Add text to the paragraph
oParagraph.AddText("This is a paragraph with the text in it aligned by the center. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.");

// Set justification to center
oParagraph.SetJc("center");

// Create a new paragraph
oParagraph = Api.CreateParagraph();

// Add text to the paragraph
oParagraph.AddText("This is a paragraph with the text in it aligned by the right side. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.");

// Set justification to right
oParagraph.SetJc("right");

// Push the paragraph to the document content
oDocContent.Push(oParagraph);

// Create another new paragraph
oParagraph = Api.CreateParagraph();

// Add text to the paragraph
oParagraph.AddText("This is a paragraph with the text in it aligned by the left side. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.");

// Set justification to left
oParagraph.SetJc("left");

// Push the paragraph to the document content
oDocContent.Push(oParagraph);
```