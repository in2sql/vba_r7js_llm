**Description / Описание:**
This code adds a shape to the active worksheet, sets its fill and stroke properties, inserts a centered paragraph with specified text, and displays the justification style used.  
Этот код добавляет фигуру на активный лист, устанавливает свойства заливки и обводки, вставляет центрированный абзац с указанным текстом и отображает используемый стиль выравнивания.

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with width 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape of type 'flowChartOnlineStorage' with specified dimensions and styles
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
);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph in the shape's content
var oParagraph = oDocContent.GetElement(0);

// Get the paragraph properties
var oParaPr = oParagraph.GetParaPr();

// Set the justification to center
oParaPr.SetJc("center");

// Add text to the paragraph
oParagraph.AddText("This is a paragraph with the text in it aligned by the center. ");
oParagraph.AddText("The justification is specified in the paragraph style. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.");

// Get the current justification setting
var sJc = oParaPr.GetJc();

// Create a new paragraph
oParagraph = Api.CreateParagraph();

// Add text displaying the justification setting
oParagraph.AddText("Justification: " + sJc);

// Add the new paragraph to the shape's content
oDocContent.Push(oParagraph);
```

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ActiveSheet

' Add a shape of type flowChartOnlineStorage with specified dimensions
Dim oShape As Shape
Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowChartOnlineStorage, _
    120, 70, 200, 150) ' Dimensions are in points

' Set the fill color to RGB (255, 111, 61)
With oShape.Fill
    .Visible = msoTrue
    .Solid
    .ForeColor.RGB = RGB(255, 111, 61)
End With

' Remove the line (stroke) from the shape
With oShape.Line
    .Visible = msoFalse
End With

' Add text to the shape
With oShape.TextFrame2
    .HorizontalAnchor = msoAnchorCenter ' Center alignment
    .TextRange.Text = "This is a paragraph with the text in it aligned by the center." & vbCrLf & _
                     "The justification is specified in the paragraph style." & vbCrLf & _
                     "These sentences are used to add lines for demonstrative purposes." & vbCrLf & _
                     "These sentences are used to add lines for demonstrative purposes." & vbCrLf & _
                     "These sentences are used to add lines for demonstrative purposes."
    
    ' Display the justification setting
    .TextRange.InsertAfter vbCrLf & "Justification: Center"
End With
```