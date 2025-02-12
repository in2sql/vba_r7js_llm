**English:** This example demonstrates how to add a shape to the active worksheet, insert two paragraphs of text into it, and set the first paragraph to bold.

**Russian:** Этот пример демонстрирует, как добавить фигуру на активный лист, вставить в нее два абзаца текста и сделать первый абзац жирным.

```javascript
// This example shows how to add a shape to the active worksheet, insert two paragraphs of text, and set the first paragraph to bold.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
// Add a shape to the worksheet with specified parameters
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
var oDocContent = oShape.GetContent(); // Get the content of the shape
oDocContent.RemoveAllElements(); // Remove all existing elements
var oParagraph1 = Api.CreateParagraph(); // Create the first paragraph
oParagraph1.AddText("This is the first paragraph."); // Add text to the first paragraph
oDocContent.Push(oParagraph1); // Add the first paragraph to the content
var oParagraph2 = Api.CreateParagraph(); // Create the second paragraph
oParagraph2.AddText("This is the second paragraph."); // Add text to the second paragraph
oDocContent.Push(oParagraph2); // Add the second paragraph to the content
var oPreviousParagraph = oParagraph2.GetPrevious(); // Get the previous paragraph relative to the second paragraph
oPreviousParagraph.SetBold(true); // Set the first paragraph to bold
```

```vba
' This example demonstrates how to add a shape to the active worksheet, insert two paragraphs of text into it, and set the first paragraph to bold.
Sub AddShapeWithParagraphs()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' Get the active worksheet
    
    ' Define fill color (RGB: 255, 111, 61)
    Dim fillColor As Long
    fillColor = RGB(255, 111, 61)
    
    ' Add a rectangle shape to the worksheet
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartInternalStorage, 60, 35, 200, 150) ' Positions and size are in points
    
    ' Set the fill color of the shape
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = fillColor
        .Solid
    End With
    
    ' Remove the line (stroke) from the shape
    With shp.Line
        .Visible = msoFalse
    End With
    
    ' Access the text frame of the shape
    With shp.TextFrame
        .Characters.Text = "" ' Remove all existing text
        
        ' Add first paragraph
        .Characters.Text = "This is the first paragraph." & vbCrLf
        ' Add second paragraph
        .Characters.Text = .Characters.Text & "This is the second paragraph."
        
        ' Set the first paragraph to bold
        .Characters(Start:=1, Length:=Len("This is the first paragraph.")).Font.Bold = True
    End With
End Sub
```