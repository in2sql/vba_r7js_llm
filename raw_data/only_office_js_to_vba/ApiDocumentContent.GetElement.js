# Description / Описание

This code adds a flowchart shape to the active sheet, sets its fill color, stroke, and positions it. It also adds a centered paragraph with multiple text lines.  
Этот код добавляет фигуру блок-схемы на активный лист, устанавливает её цвет заливки, обводку и позиционирует её. Также добавляется выровненный по центру абзац с несколькими строками текста.

```vba
' VBA Code
Sub AddFlowChartShape()
    Dim oWorksheet As Worksheet
    Dim oShape As Shape
    Dim oFill As Long
    Dim oStroke As Long
    
    ' Set the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Define fill color (RGB)
    oFill = RGB(255, 111, 61)
    
    ' Add a flowchart shape with specific dimensions
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, 200, 60, 200, 60)
    
    ' Set the fill color
    With oShape.Fill
        .ForeColor.RGB = oFill
        .Visible = msoTrue
    End With
    
    ' Remove the stroke
    With oShape.Line
        .Visible = msoFalse
    End With
    
    ' Add and format the text within the shape
    With oShape.TextFrame2.TextRange
        .ParagraphFormat.Alignment = msoAlignCenter ' Center alignment
        .Text = "This is a paragraph with the text in it aligned by the center." & vbCrLf & _
                "The justification is specified in the paragraph style." & vbCrLf & _
                "These sentences are used to add lines for demonstrative purposes." & vbCrLf & _
                "These sentences are used to add lines for demonstrative purposes."
    End With
End Sub
```

```javascript
// OnlyOffice JS Code
// This example shows how to get an element by its position in the document content.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 200 * 36000, 60 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000); // Add the shape
var oDocContent = oShape.GetContent(); // Get the content of the shape
var oParagraph = oDocContent.GetElement(0); // Get the first paragraph
var oParaPr = oParagraph.GetParaPr(); // Get paragraph properties
oParaPr.SetJc("center"); // Set justification to center
oParagraph.AddText("This is a paragraph with the text in it aligned by the center. "); // Add text
oParagraph.AddText("The justification is specified in the paragraph style. "); // Add text
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. "); // Add text
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes."); // Add text
```