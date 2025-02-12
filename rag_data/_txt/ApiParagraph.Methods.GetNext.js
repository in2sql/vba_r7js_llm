**Description / Описание**

This code demonstrates how to manipulate shapes and paragraphs in OnlyOffice using JavaScript and their equivalent implementation in Excel VBA.

Этот код демонстрирует, как манипулировать фигурами и абзацами в OnlyOffice с использованием JavaScript и их эквивалентную реализацию в Excel VBA.

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with width 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified parameters
var oShape = oWorksheet.AddShape(
    "flowChartOnlineStorage", 
    60 * 36000, 
    35 * 36000, 
    oFill, 
    oStroke, 
    0, 
    2 * 36000, 
    0, 
    3 * 36000
);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Remove all existing elements from the content
oDocContent.RemoveAllElements();

// Create the first paragraph and add text
var oParagraph1 = Api.CreateParagraph();
oParagraph1.AddText("This is the first paragraph.");
oDocContent.Push(oParagraph1);

// Create the second paragraph and add text
var oParagraph2 = Api.CreateParagraph();
oParagraph2.AddText("This is the second paragraph.");
oDocContent.Push(oParagraph2);

// Get the next paragraph after the first and set it to bold
var oNextParagraph = oParagraph1.GetNext();
oNextParagraph.SetBold(true);
```

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ActiveSheet

' Create a solid fill with RGB color (255, 111, 61)
Dim oFill As ShapeFill
Set oFill = oWorksheet.Shapes.Fill
oFill.ForeColor.RGB = RGB(255, 111, 61)
oFill.Solid

' Create a stroke with width 0 and no fill
Dim oStroke As ShapeLineFormat
Set oStroke = oWorksheet.Shapes.Line
With oStroke
    .Weight = 0
    .Visible = msoFalse
End With

' Add a shape to the worksheet with specified parameters
Dim oShape As Shape
Set oShape = oWorksheet.Shapes.AddShape( _
    Type:=msoShapeFlowchartInternalStorage, _
    Left:=60 * 72, _ ' Assuming 72 DPI conversion
    Top:=35 * 72, _
    Width:=2 * 72, _
    Height:=3 * 72 _
)
With oShape
    .Fill.ForeColor.RGB = RGB(255, 111, 61)
    .Line.Visible = msoFalse
End With

' Get the text frame of the shape
Dim oTextFrame As TextFrame
Set oTextFrame = oShape.TextFrame

' Remove all existing text
oTextFrame.TextRange.Text = ""

' Create the first paragraph and add text
Dim oParagraph1 As TextRange
Set oParagraph1 = oTextFrame.TextRange.Paragraphs(1)
oParagraph1.Text = "This is the first paragraph."

' Create the second paragraph and add text
Dim oParagraph2 As TextRange
Set oParagraph2 = oTextFrame.TextRange.Paragraphs.Add
oParagraph2.Text = "This is the second paragraph."

' Get the next paragraph after the first and set it to bold
With oParagraph2.Font
    .Bold = True
End With
```