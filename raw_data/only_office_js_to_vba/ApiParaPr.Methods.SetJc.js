# Description / Описание

**English:**

This code sets the justification of the paragraph contents to center within a shape in the active worksheet. It creates a fill and stroke for the shape, adds the shape to the worksheet, retrieves the paragraph properties, sets the alignment to center, and adds multiple lines of text to the paragraph.

**Russian:**

Этот код устанавливает выравнивание содержимого абзаца по центру внутри формы в активном листе. Он создает заливку и обводку для формы, добавляет форму на лист, получает свойства абзаца, устанавливает выравнивание по центру и добавляет несколько строк текста в абзац.

```vba
' VBA code equivalent

Sub SetParagraphJustification()
    ' Get the active sheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Add a shape to the worksheet
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(Type:=msoShapeFlowchartOnlineStorage, _
        Left:=120 * 36000, Top:=70 * 36000, Width:=200, Height:=100)
    
    ' Set the fill color to RGB(255, 111, 61)
    oShape.Fill.ForeColor.RGB = RGB(255, 111, 61)
    
    ' Remove the shape's outline
    oShape.Line.Visible = msoFalse
    
    ' Access the text frame of the shape
    With oShape.TextFrame2
        ' Set the paragraph alignment to center
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        
        ' Add text to the paragraph
        .TextRange.Text = "This is a paragraph with the text in it aligned by the center. " & _
                          "The justification is specified in the paragraph style. " & _
                          "These sentences are used to add lines for demonstrative purposes. " & _
                          "These sentences are used to add lines for demonstrative purposes."
    End With
End Sub
```

```javascript
// OnlyOffice JS code equivalent

// Get the active sheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with width 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified position and size
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Get the paragraph properties
var oParaPr = oParagraph.GetParaPr();

// Set the alignment to center
oParaPr.SetJc("center");

// Add text to the paragraph
oParagraph.AddText("This is a paragraph with the text in it aligned by the center. ");
oParagraph.AddText("The justification is specified in the paragraph style. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.");
```