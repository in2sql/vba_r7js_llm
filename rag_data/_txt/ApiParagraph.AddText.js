**Description / Описание:**

English: This code adds a shape to the active sheet with specified fill and stroke, and inserts aligned text with a line break inside the shape.

Russian: Этот код добавляет фигуру на активный лист с указанной заливкой и контуром, а также вставляет выровненный текст с разрывом строки внутри фигуры.

---

```vba
' VBA code equivalent to OnlyOffice JS example

Sub AddShapeWithText()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Define fill color using RGB
    Dim fillColor As Long
    fillColor = RGB(255, 111, 61)
    
    ' Add a shape to the worksheet
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape( _
        msoShapeFlowchartOfflineStorage, _ ' Shape type
        120, _ ' Left position in points
        70, _  ' Top position in points
        200, _ ' Width in points
        150)  ' Height in points
    
    ' Apply solid fill color to the shape
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = fillColor
        .Solid
    End With
    
    ' Remove the shape's stroke
    With oShape.Line
        .Visible = msoFalse
    End With
    
    ' Add and format text inside the shape
    With oShape.TextFrame2
        .TextRange.ParagraphFormat.Alignment = msoAlignLeft ' Align text to the left
        .TextRange.Text = "This is a text inside the shape aligned left." & vbCrLf & "This is a text after the line break."
    End With
End Sub
```

---

```javascript
// OnlyOffice JS code equivalent to the VBA example

// Get the active sheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet
var oShape = oWorksheet.AddShape(
    "flowChartOnlineStorage",            // Shape type
    120 * 36000,                         // Left position in EMUs
    70 * 36000,                          // Top position in EMUs
    oFill,                               // Fill
    oStroke,                             // Stroke
    0,                                   // Rotation
    2 * 36000,                           // Width in EMUs
    0,                                   // Height in EMUs
    3 * 36000                            // Another dimension if needed
);

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