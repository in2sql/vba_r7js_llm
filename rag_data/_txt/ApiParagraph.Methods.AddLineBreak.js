**Description / Описание**

This code adds a shape to the active worksheet, fills it with a specified color, removes the stroke, adds text aligned to the left, inserts a line break, and adds another line of text.

Этот код добавляет фигуру на активный лист, заполняет её указанным цветом, удаляет обводку, добавляет текст с выравниванием по левому краю, вставляет разрыв строки и добавляет ещё одну строку текста.

```vba
' VBA Code to add a shape with text and line break in Excel

Sub AddShapeWithText()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Create a solid fill with RGB color (255, 111, 61)
    Dim fillColor As Long
    fillColor = RGB(255, 111, 61)
    
    ' Add the shape to the worksheet
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartPredefined, _
                                 100, 100, 300, 150) ' Adjust position and size as needed
                                 
    ' Apply fill color
    shp.Fill.ForeColor.RGB = fillColor
    
    ' Remove the stroke
    shp.Line.Visible = msoFalse
    
    ' Add text to the shape
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
    shp.TextFrame2.TextRange.Text = "This is a text inside the shape aligned left." & vbCrLf & "This is a text after the line break."
End Sub
```

```javascript
// JavaScript Code using OnlyOffice API to add a shape with text and line break

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Set paragraph alignment to left
oParagraph.SetJc("left");

// Add text to the paragraph
oParagraph.AddText("This is a text inside the shape aligned left.");

// Add a line break
oParagraph.AddLineBreak();

// Add more text after the line break
oParagraph.AddText("This is a text after the line break.");
```