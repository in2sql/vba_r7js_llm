**Description:**

*English:* This script adds a shape to the active worksheet, sets its fill color and stroke, and modifies the paragraph spacing within the shape's content. It adds two paragraphs of text with specific spacing after the first paragraph.

*Russian:* Этот скрипт добавляет фигуру на активный лист, устанавливает цвет заливки и обводки, а также изменяет межстрочный интервал в содержимом фигуры. Он добавляет два абзаца текста с определённым отступом после первого абзаца.

**JavaScript Code:**

```javascript
// This example sets the spacing after the current paragraph.

// Get the active sheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with width 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified parameters
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape's document
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Get paragraph properties
var oParaPr = oParagraph.GetParaPr();

// Set spacing after the paragraph
oParaPr.SetSpacingAfter(1440);

// Add text to the paragraph
oParagraph.AddText("This is an example of setting a space after a paragraph. ");
oParagraph.AddText("The second paragraph will have an offset of one inch from the top. ");
oParagraph.AddText("This is due to the fact that the first paragraph has this offset enabled.");

// Create a new paragraph
oParagraph = Api.CreateParagraph();
oParagraph.AddText("This is the second paragraph and it is one inch away from the first paragraph.");

// Push the new paragraph to the document content
oDocContent.Push(oParagraph);
```

**VBA Code:**

```vba
' This example adds a shape to the active worksheet, sets its fill color and stroke,
' and modifies the paragraph spacing within the shape's text.

Sub AddShapeWithParagraphSpacing()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Add a shape to the worksheet
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartOnlineStorage, 120, 70, 200, 100)
    
    ' Set the fill color to RGB(255, 111, 61)
    shp.Fill.ForeColor.RGB = RGB(255, 111, 61)
    
    ' Set no line (stroke)
    shp.Line.Visible = msoFalse
    
    ' Add text to the shape
    shp.TextFrame2.TextRange.Text = "This is an example of setting a space after a paragraph." & vbCrLf & _
        "The second paragraph will have an offset of one inch from the top." & vbCrLf & _
        "This is due to the fact that the first paragraph has this offset enabled." & vbCrLf & _
        "This is the second paragraph and it is one inch away from the first paragraph."
    
    ' Modify paragraph spacing
    Dim para As TextRange2
    Set para = shp.TextFrame2.TextRange.Paragraphs(1)
    para.SpaceAfter = 1440 ' Points
    
    ' Optionally, set spacing before for the second paragraph if needed
    Set para = shp.TextFrame2.TextRange.Paragraphs(2)
    para.SpaceBefore = 1440 ' Points
End Sub
```