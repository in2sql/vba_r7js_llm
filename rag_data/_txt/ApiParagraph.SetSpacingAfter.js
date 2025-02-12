# Description / Описание

**English:** This script creates a shape within the active Excel worksheet, sets its fill color, removes the stroke, adds multiple paragraphs of text, and adjusts the spacing after the first paragraph.

**Russian:** Этот скрипт создает фигуру на активном листе Excel, устанавливает ее цвет заливки, удаляет обводку, добавляет несколько абзацев текста и настраивает отступ после первого абзаца.

## VBA Code:

```vba
' This VBA script creates a shape, sets its fill color, removes the stroke,
' adds multiple paragraphs of text, and adjusts the spacing after the first paragraph.

Sub CreateShapeWithText()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim txtFrame As TextFrame2
    Dim paragraph As TextRange2
    
    ' Get the active worksheet
    Set ws = ActiveSheet
    
    ' Add a shape to the worksheet
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartStoredData, _
                                 120, 70, 200, 100)
    
    ' Set the fill color to RGB(255, 111, 61)
    shp.Fill.ForeColor.RGB = RGB(255, 111, 61)
    
    ' Remove the stroke
    shp.Line.Visible = msoFalse
    
    ' Access the text frame of the shape
    Set txtFrame = shp.TextFrame2
    
    ' Add text to the shape
    txtFrame.TextRange.Text = "This is an example of setting a space after a paragraph." & vbCrLf & _
                              "The second paragraph will have an offset of one inch from the top." & vbCrLf & _
                              "This is due to the fact that the first paragraph has this offset enabled."
    
    ' Set spacing after the first paragraph
    Set paragraph = txtFrame.TextRange.Paragraphs(1)
    paragraph.SpaceAfter = 12 ' Points (approx. 1 inch)
    
    ' Add a second paragraph with spacing
    txtFrame.TextRange.Text = txtFrame.TextRange.Text & vbCrLf & "This is the second paragraph and it is one inch away from the first paragraph."
End Sub
```

## OnlyOffice JS Code:

```javascript
// This example sets the spacing after the paragraph.

var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified dimensions and styles
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 
                                 120 * 36000, 
                                 70 * 36000, 
                                 oFill, 
                                 oStroke, 
                                 0, 
                                 2 * 36000, 
                                 0, 
                                 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph in the content
var oParagraph = oDocContent.GetElement(0);

// Add text to the first paragraph
oParagraph.AddText("This is an example of setting a space after a paragraph. ");
oParagraph.AddText("The second paragraph will have an offset of one inch from the top. ");
oParagraph.AddText("This is due to the fact that the first paragraph has this offset enabled.");

// Set spacing after the first paragraph
oParagraph.SetSpacingAfter(1440);

// Create a new paragraph
oParagraph = Api.CreateParagraph();

// Add text to the second paragraph
oParagraph.AddText("This is the second paragraph and it is one inch away from the first paragraph.");

// Add the second paragraph to the document content
oDocContent.Push(oParagraph);
```