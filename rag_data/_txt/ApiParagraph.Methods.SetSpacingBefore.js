# Description

**English:**  
Adds a flowchart shape to the active worksheet with specified fill and stroke colors. Within the shape, it adds two paragraphs of text, setting spacing before the second paragraph.

**Russian:**  
Добавляет фигуру блок-схемы на активный лист с указанными цветами заливки и обводки. Внутри фигуры добавляет два абзаца текста, устанавливая отступ перед вторым абзацем.

## VBA Code

```vba
' Adds a flowchart shape to the active worksheet with specified fill and stroke colors
' and adds two paragraphs of text with spacing before the second paragraph.

Sub AddFlowchartShape()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim fillColor As Long
    Dim lineColor As Long
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Define fill color using RGB
    fillColor = RGB(255, 111, 61)
    
    ' Define no line color (using transparent)
    lineColor = RGB(255, 255, 255) ' White color as placeholder
    
    ' Add the flowchart shape (using a predefined flowchart type)
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartProcess, 120, 70, 120, 70)
    
    ' Set the fill color
    With shp.Fill
        .Visible = msoTrue
        .Solid
        .ForeColor.RGB = fillColor
    End With
    
    ' Set the line color and weight
    With shp.Line
        .Visible = msoTrue
        .ForeColor.RGB = lineColor
        .Weight = 0 ' No border
    End With
    
    ' Add text to the shape
    With shp.TextFrame2
        .TextRange.Text = "This is an example of setting a space before the paragraph." & vbCrLf & _
                          "The second paragraph will have an offset of one inch from the top." & vbCrLf & _
                          "This is due to the fact that the second paragraph has this offset enabled."
        ' Set paragraph spacing before for the second paragraph
        .TextRange.Paragraphs(2).ParagraphFormat.SpaceBefore = 144 ' Points (approx. one inch)
    End With
End Sub
```

## OnlyOffice JS Code

```javascript
// Adds a flowchart shape to the active worksheet with specified fill and stroke colors
// and adds two paragraphs of text with spacing before the second paragraph.

function addFlowchartShape() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Create a solid fill with RGB color (255, 111, 61)
    var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
    
    // Create a stroke with no fill
    var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    
    // Add the flowchart shape with specified dimensions and styling
    var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
    
    // Get the content of the shape
    var oDocContent = oShape.GetContent();
    
    // Get the first paragraph and add text
    var oParagraph = oDocContent.GetElement(0);
    oParagraph.AddText("This is an example of setting a space before a paragraph. ");
    oParagraph.AddText("The second paragraph will have an offset of one inch from the top. ");
    oParagraph.AddText("This is due to the fact that the second paragraph has this offset enabled.");
    
    // Create a second paragraph
    oParagraph = Api.CreateParagraph();
    oParagraph.AddText("This is the second paragraph and it is one inch away from the first paragraph.");
    
    // Set spacing before the second paragraph (1440 units)
    oParagraph.SetSpacingBefore(1440);
    
    // Add the second paragraph to the document content
    oDocContent.Push(oParagraph); 
}
```