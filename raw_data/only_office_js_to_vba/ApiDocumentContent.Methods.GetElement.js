## Description / Описание

**English:**  
This code adds a flowchart shape to the active worksheet, sets its fill and stroke properties, centers the paragraph within the shape, and adds multiple lines of centered text.

**Russian:**  
Этот код добавляет форму блок-схемы на активный лист, устанавливает свойства заливки и обводки, центрирует параграф внутри формы и добавляет несколько строк центрированного текста.

---

### OnlyOffice JavaScript Code

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with width 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a flowchart shape to the worksheet with specified dimensions and styles
var oShape = oWorksheet.AddShape(
    "flowChartOnlineStorage", // Shape type
    200 * 36000,              // Left position
    60 * 36000,               // Top position
    oFill,                    // Fill style
    oStroke,                  // Stroke style
    0,                        // Rotation
    2 * 36000,                // Width
    0,                        // Height
    3 * 36000                 // Additional parameter (if any)
);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph in the shape's content
var oParagraph = oDocContent.GetElement(0);

// Get the paragraph properties
var oParaPr = oParagraph.GetParaPr();

// Set the paragraph alignment to center
oParaPr.SetJc("center");

// Add multiple lines of text to the paragraph
oParagraph.AddText("This is a paragraph with the text in it aligned by the center. ");
oParagraph.AddText("The justification is specified in the paragraph style. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.");
```

---

### Excel VBA Code

```vba
Sub AddFlowChartShape()
    ' Get the active worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Define the RGB color
    Dim fillColor As Long
    fillColor = RGB(255, 111, 61)
    
    ' Add a flowchart shape to the worksheet
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape( _
        Type:=msoShapeFlowchartProcess, _ ' Equivalent to "flowChartOnlineStorage"
        Left:=200, _                      ' Position from left (in points)
        Top:=60, _                        ' Position from top (in points)
        Width:=200, _                     ' Width of the shape
        Height:=60                        ' Height of the shape
    )
    
    ' Set the fill color
    With shp.Fill
        .Visible = msoTrue
        .Solid
        .ForeColor.RGB = fillColor
    End With
    
    ' Set the line properties (no line)
    With shp.Line
        .Visible = msoFalse
    End With
    
    ' Access the text frame within the shape
    With shp.TextFrame2
        ' Center the text horizontally
        .HorizontalAnchor = msoAnchorCenter
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        
        ' Add multiple lines of text
        .TextRange.Text = "This is a paragraph with the text in it aligned by the center." & vbCrLf & _
                          "The justification is specified in the paragraph style." & vbCrLf & _
                          "These sentences are used to add lines for demonstrative purposes." & vbCrLf & _
                          "These sentences are used to add lines for demonstrative purposes."
    End With
End Sub
```