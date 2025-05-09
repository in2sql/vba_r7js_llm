# Description / Описание

**English:**  
This code adds a flowchart shape to the active worksheet, sets its fill color and stroke, inserts text into its paragraphs, and adjusts the spacing before the second paragraph.

**Russian:**  
Этот код добавляет форму блок-схемы на активный лист, устанавливает цвет заливки и обводку, вставляет текст в абзацы и регулирует отступ перед вторым абзацем.

```vba
' VBA Code

Sub AddFlowchartShape()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Define position and size (units in points)
    Dim LeftPos As Single, TopPos As Single, Width As Single, Height As Single
    LeftPos = 120 * 10 ' Example conversion from original units
    TopPos = 70 * 10
    Width = 200
    Height = 150
    
    ' Add a flowchart process shape to the worksheet
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartProcess, LeftPos, TopPos, Width, Height)
    
    ' Set the fill color to RGB(255, 111, 61)
    oShape.Fill.ForeColor.RGB = RGB(255, 111, 61)
    
    ' Remove the stroke
    oShape.Line.Visible = msoFalse
    
    ' Add text to the shape
    With oShape.TextFrame2.TextRange
        ' Add first paragraph
        .Text = "This is an example of setting a space before the paragraph. " & _
                "The second paragraph will have an offset of one inch from the top. " & _
                "This is due to the fact that the second paragraph has this offset enabled." & vbCrLf
        ' Add second paragraph
        .InsertAfter "This is the second paragraph and it is one inch away from the first paragraph."
    End With
    
    ' Set spacing before for the second paragraph
    With oShape.TextFrame2.TextRange.Paragraphs(2).ParagraphFormat
        .SpaceBefore = 14.4 ' Points (1 inch = 72 points)
    End With
End Sub
```

```javascript
// JavaScript Code

// This example sets the spacing before the paragraph.
// Этот пример устанавливает отступ перед абзацем.

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create fill color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create stroke with no fill

// Add a flowchart shape to the worksheet with specified dimensions and fill/stroke
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph
var oParagraph = oDocContent.GetElement(0);

// Add text to the first paragraph
oParagraph.AddText("This is an example of setting a space before a paragraph. ");
oParagraph.AddText("The second paragraph will have an offset of one inch from the top. ");
oParagraph.AddText("This is due to the fact that the second paragraph has this offset enabled.");

// Create a new paragraph
oParagraph = Api.CreateParagraph();

// Add text to the second paragraph
oParagraph.AddText("This is the second paragraph and it is one inch away from the first paragraph.");

// Set spacing before the second paragraph
oParagraph.SetSpacingBefore(1440); // 1440 units for spacing

// Add the new paragraph to the document content
oDocContent.Push(oParagraph);
```