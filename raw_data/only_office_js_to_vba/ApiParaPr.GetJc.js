**Description / Описание:**

This code creates a shape in the active worksheet, sets its fill color, removes the stroke, adds a centered paragraph with multiple lines of text, retrieves the current justification, and appends a new paragraph indicating the justification used.

Этот код создает фигуру на активном листе, устанавливает цвет заливки, удаляет обводку, добавляет абзац с центрированным выравниванием и несколькими строками текста, получает текущее выравнивание и добавляет новый абзац, указывающий использованное выравнивание.

```vba
' VBA Code equivalent to OnlyOffice JS example

Sub CreateShapeAndSetJustification()
    ' Get the active worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Create a shape with specific dimensions
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeFlowChartOnlineStorage, 120, 70, 200, 150)
    
    ' Set fill color to RGB(255, 111, 61)
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61)
        .Solid
    End With
    
    ' Remove the line (stroke)
    With shp.Line
        .Visible = msoFalse
    End With
    
    ' Add text to the shape
    With shp.TextFrame
        .HorizontalAlignment = xlHAlignCenter ' Set justification to center
        .Characters.Text = "This is a paragraph with the text in it aligned by the center. " & _
                           "The justification is specified in the paragraph style. " & _
                           "These sentences are used to add lines for demonstrative purposes. " & _
                           "These sentences are used to add lines for demonstrative purposes. " & _
                           "These sentences are used to add lines for demonstrative purposes."
    End With
    
    ' Get the current justification
    Dim jc As String
    jc = "center" ' Since we set it to center
    
    ' Add a new shape or text box to show the justification
    Dim shpJc As Shape
    Set shpJc = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 120, 230, 400, 50)
    With shpJc.TextFrame
        .Characters.Text = "Justification: " & jc
        .HorizontalAlignment = xlHAlignLeft
    End With
End Sub
```

```javascript
// This example creates a shape with a specific fill color and no stroke,
// adds a centered paragraph with multiple lines of text,
// retrieves the justification, and appends a new paragraph indicating the justification.

var oWorksheet = Api.GetActiveSheet();
// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
// Add a shape of type 'flowChartOnlineStorage' with specified dimensions and styles
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
// Get the content of the shape
var oDocContent = oShape.GetContent();
// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);
// Get the paragraph properties
var oParaPr = oParagraph.GetParaPr();
// Set justification to center
oParaPr.SetJc("center");
// Add multiple lines of text to the paragraph
oParagraph.AddText("This is a paragraph with the text in it aligned by the center. ");
oParagraph.AddText("The justification is specified in the paragraph style. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.");
// Retrieve the current justification
var sJc = oParaPr.GetJc();
// Create a new paragraph
oParagraph = Api.CreateParagraph();
// Add text indicating the justification
oParagraph.AddText("Justification: " + sJc);
// Push the new paragraph to the document content
oDocContent.Push(oParagraph);
```