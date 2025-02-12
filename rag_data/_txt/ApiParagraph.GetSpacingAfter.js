## Description
This code adds a flowchart shape to the active worksheet, inserts two paragraphs of text into the shape with specific spacing after the first paragraph, and retrieves the spacing value.

## Описание
Этот код добавляет форму блок-схемы на активный лист, вставляет два абзаца текста в форму с определенным отступом после первого абзаца и извлекает значение отступа.

### VBA Code
```vba
' Add a flowchart shape to the active worksheet, insert two paragraphs with spacing, and retrieve spacing after

Sub AddFlowChartWithParagraphs()
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Create a solid fill with RGB color (255, 111, 61)
    Dim oFill As Object
    Set oFill = ActiveWorkbook.Styles.Add("CustomFill")
    oFill.Interior.Color = RGB(255, 111, 61)
    
    ' Create no line (stroke)
    Dim oStroke As Object
    Set oStroke = Nothing ' No fill for the stroke
    
    ' Add the shape to the worksheet
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowChartPredefined, 120, 70, 300, 150)
    oShape.Fill.ForeColor.RGB = RGB(255, 111, 61)
    oShape.Line.Visible = msoFalse
    
    ' Add text to the shape
    With oShape.TextFrame2.TextRange
        .Text = "This is an example of setting a space after a paragraph." & vbCrLf & _
                "The second paragraph will have an offset of one inch from the top." & vbCrLf & _
                "This is due to the fact that the first paragraph has this offset enabled."
                
        ' Set spacing after first paragraph
        .Paragraphs(1).ParagraphFormat.SpaceAfter = 12 ' Points
        
        ' Add second paragraph
        .Text = .Text & vbCrLf & "This is the second paragraph and it is one inch away from the first paragraph." & vbCrLf
        .Paragraphs(2).ParagraphFormat.SpaceAfter = .Paragraphs(1).ParagraphFormat.SpaceAfter
        
        ' Retrieve spacing after from the first paragraph
        Dim nSpacingAfter As Single
        nSpacingAfter = .Paragraphs(1).ParagraphFormat.SpaceAfter
        .Text = .Text & "Spacing after: " & nSpacingAfter & " pts"
    End With
End Sub
```

### JS Code
```javascript
// Adds a flowchart shape to the active sheet, inserts two paragraphs with spacing, and retrieves spacing after
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create no stroke
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add the shape to the worksheet with specified dimensions and formatting
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph and add text
var oParagraph1 = oDocContent.GetElement(0);
oParagraph1.AddText("This is an example of setting a space after a paragraph. ");
oParagraph1.AddText("The second paragraph will have an offset of one inch from the top. ");
oParagraph1.AddText("This is due to the fact that the first paragraph has this offset enabled.");

// Set spacing after the first paragraph
oParagraph1.SetSpacingAfter(1440);

// Create the second paragraph and add text
var oParagraph2 = Api.CreateParagraph();
oParagraph2.AddText("This is the second paragraph and it is one inch away from the first paragraph.");
oParagraph2.AddLineBreak();

// Retrieve spacing after from the first paragraph
var nSpacingAfter = oParagraph1.GetSpacingAfter();
oParagraph2.AddText("Spacing after: " + nSpacingAfter);

// Add the second paragraph to the document content
oDocContent.Push(oParagraph2);
```