---
**Description:**
This code adds a shape to the active worksheet, sets its fill and stroke properties, modifies the shape's text content by adding multiple paragraphs, and removes a specific paragraph element.

**Описание:**
Этот код добавляет фигуру на активный лист, устанавливает свойства заливки и обводки, изменяет текстовое содержимое фигуры, добавляя несколько параграфов, и удаляет определенный элемент параграфа.
---

```javascript
// JavaScript code to add and modify a shape in OnlyOffice

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with width 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape with specified type, dimensions, fill, stroke, and position
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Remove all existing elements from the paragraph
oParagraph.RemoveAllElements();

// Create and add the first run of text to the paragraph
var oRun = Api.CreateRun();
oRun.AddText("This is the first paragraph element. ");
oParagraph.AddElement(oRun);

// Create and add the second run of text to the paragraph
oRun = Api.CreateRun();
oRun.AddText("This is the second paragraph element. ");
oParagraph.AddElement(oRun);

// Create and add the third run of text to the paragraph
oRun = Api.CreateRun();
oRun.AddText("This is the third paragraph element (it will be removed from the paragraph and we will not see it). ");
oParagraph.AddElement(oRun);

// Add a line break to the paragraph
oParagraph.AddLineBreak();

// Create and add the fourth run of text to the paragraph
oRun = Api.CreateRun();
oRun.AddText("This is the fourth paragraph element - it became the third, because we removed the previous run from the paragraph. ");
oParagraph.AddElement(oRun);

// Add a line break to the paragraph
oParagraph.AddLineBreak();

// Create and add the fifth run of text to the paragraph
oRun = Api.CreateRun();
oRun.AddText("Please note that line breaks are not counted into paragraph elements!");
oParagraph.AddElement(oRun);

// Remove the third element from the paragraph
oParagraph.RemoveElement(3);
```

```vba
' VBA code to add and modify a shape in Excel

Sub ModifyShape()
    ' Get the active worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Create RGB color (255, 111, 61)
    Dim fillColor As Long
    fillColor = RGB(255, 111, 61)
    
    ' Add a shape with specified type and position
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartProcess, 120, 70, 200, 100) ' Adjust width and height as needed
    
    ' Set the fill properties
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = fillColor
        .Solid
    End With
    
    ' Set the line (stroke) properties
    With shp.Line
        .Weight = 0
        .Visible = msoFalse
    End With
    
    ' Clear existing text
    shp.TextFrame2.TextRange.Text = ""
    
    ' Add text runs with line breaks
    With shp.TextFrame2.TextRange
        .Text = "This is the first paragraph element. " & vbCrLf & _
                "This is the second paragraph element. " & vbCrLf & _
                "This is the third paragraph element (it will be removed from the paragraph and we will not see it). " & vbCrLf & _
                "This is the fourth paragraph element - it became the third, because we removed the previous run from the paragraph. " & vbCrLf & _
                "Please note that line breaks are not counted into paragraph elements!"
        
        ' Remove the third paragraph
        .Paragraphs(3).Delete
    End With
End Sub
```