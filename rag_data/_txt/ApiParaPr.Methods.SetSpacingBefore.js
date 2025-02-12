### Description
This code adds a shape to the active worksheet, sets its fill and stroke properties, adds text with multiple paragraphs, and adjusts the spacing before a specific paragraph.

Этот код добавляет форму на активный лист, устанавливает ее свойства заливки и обводки, добавляет текст с несколькими абзацами и настраивает отступ перед определенным абзацем.

```vba
' VBA code to add a shape, set fill and stroke, add text with paragraphs, and set spacing before a paragraph

Sub AddShapeWithParagraphs()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Add the shape to the worksheet
    ' msoShapeFlowchartOnlineStorage corresponds to "flowChartOnlineStorage"
    ' Left and Top positions are in points; Width and Height are also in points
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, 120, 70, 120, 70)
    
    ' Set the fill color to RGB(255, 111, 61)
    With oShape.Fill
        .Visible = msoTrue
        .Solid
        .ForeColor.RGB = RGB(255, 111, 61)
    End With
    
    ' Set the stroke (line) properties: no visible stroke
    With oShape.Line
        .Visible = msoFalse
        .Weight = 0
    End With
    
    ' Add text to the shape
    With oShape.TextFrame2
        .TextRange.Text = "This is an example of setting a space before a paragraph. " & _
                          "The second paragraph will have an offset of one inch from the top. " & _
                          "This is due to the fact that the second paragraph has this offset enabled." & vbCrLf & _
                          "This is the second paragraph and it is one inch away from the first paragraph."
        
        ' Set spacing before for the second paragraph (1440 twips = 1 inch)
        .TextRange.Paragraphs(2).ParagraphFormat.SpaceBefore = 1440
    End With
End Sub
```

```javascript
// This example sets the spacing before the current paragraph.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill with RGB color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000); // Add the shape to the worksheet
var oDocContent = oShape.GetContent(); // Get the content of the shape
var oParagraph = oDocContent.GetElement(0); // Get the first paragraph
oParagraph.AddText("This is an example of setting a space before a paragraph. "); // Add text to the first paragraph
oParagraph.AddText("The second paragraph will have an offset of one inch from the top. "); // Add more text
oParagraph.AddText("This is due to the fact that the second paragraph has this offset enabled."); // Add more text
oParagraph = Api.CreateParagraph(); // Create a new paragraph
var oParaPr = oParagraph.GetParaPr(); // Get paragraph properties
oParaPr.SetSpacingBefore(1440); // Set spacing before to 1440 twips (1 inch)
oParagraph.AddText("This is the second paragraph and it is one inch away from the first paragraph."); // Add text to the second paragraph
oDocContent.Push(oParagraph); // Add the new paragraph to the document content
```