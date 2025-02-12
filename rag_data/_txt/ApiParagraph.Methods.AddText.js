**Description:**
This code adds a shape to the active sheet, sets its fill and stroke colors, and inserts text into the shape with specific alignment and line breaks.

**Описание:**
Этот код добавляет форму на активный лист, устанавливает цвета заливки и обводки, а также вставляет текст в форму с определенным выравниванием и разрывами строк.

```javascript
// This example adds some text to the paragraph.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
var oShape = oWorksheet.AddShape(
    "flowChartOnlineStorage",
    120 * 36000,
    70 * 36000,
    oFill,
    oStroke,
    0,
    2 * 36000,
    0,
    3 * 36000
); // Add a shape to the worksheet
var oDocContent = oShape.GetContent(); // Get the content of the shape
var oParagraph = oDocContent.GetElement(0); // Get the first paragraph element
oParagraph.SetJc("left"); // Set paragraph alignment to left
oParagraph.AddText("This is a text inside the shape aligned left."); // Add text to the paragraph
oParagraph.AddLineBreak(); // Add a line break
oParagraph.AddText("This is a text after the line break."); // Add text after the line break
```

```vba
' This VBA code adds a shape to the active sheet, sets its fill and line properties,
' and adds text with alignment and line breaks to the shape.

Sub AddShapeWithText()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' Get the active worksheet
    
    ' Define the fill color
    Dim fillColor As Long
    fillColor = RGB(255, 111, 61) ' Define RGB color
    
    ' Add a shape to the worksheet
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape( _
        Type:=msoShapeFlowchartOnlineStorage, _
        Left:=120 * 72, _ ' Convert 120 to points (assuming 72 points per inch)
        Top:=70 * 72, _ ' Convert 70 to points
        Width:=200, _
        Height:=100) ' Set width and height in points
    
    ' Set the fill color
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = fillColor
        .Solid
    End With
    
    ' Set the line properties
    With shp.Line
        .Visible = msoFalse ' No line stroke
    End With
    
    ' Add text to the shape
    With shp.TextFrame2.TextRange
        .ParagraphFormat.Alignment = msoAlignLeft ' Set alignment to left
        .Text = "This is a text inside the shape aligned left." & vbLf & _
                "This is a text after the line break." ' Add text with line break
    End With
End Sub
```