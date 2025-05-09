## Description / Описание

**English:**  
This code creates a shape with a specified fill color and no stroke, removes any existing content within the shape, sets the text properties to a font size of 30 and bold weight, aligns the text to the left, and adds a sample text paragraph to the shape.

**Russian:**  
Этот код создает фигуру с указанным цветом заливки и без обводки, удаляет любое существующее содержимое внутри фигуры, устанавливает свойства текста с размером шрифта 30 и полужирным начертанием, выравнивает текст по левому краю и добавляет пример абзаца текста в фигуру.

---

### VBA Code

```vba
' This VBA macro creates a shape with specified fill and no stroke,
' removes existing content, sets text properties, and adds a paragraph.

Sub CreateCustomShape()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Create a solid fill color (RGB: 255, 111, 61)
    Dim fillColor As Long
    fillColor = RGB(255, 111, 61)
    
    ' Add a rectangle shape
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartProcess, 100, 50, 200, 100)
    
    ' Set fill color
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = fillColor
        .Transparency = 0
    End With
    
    ' Remove the stroke (no line)
    With shp.Line
        .Visible = msoFalse
    End With
    
    ' Clear any existing text
    shp.TextFrame.Characters.Text = ""
    
    ' Set text properties
    With shp.TextFrame.Characters.Font
        .Size = 30
        .Bold = True
    End With
    
    ' Set paragraph alignment to left
    shp.TextFrame.HorizontalAlignment = xlHAlignLeft
    
    ' Add sample text
    shp.TextFrame.Characters.Text = "This is a sample text with the font size set to 30 and the font weight set to bold."
End Sub
```

### JavaScript Code

```javascript
// This example creates a shape with specified fill and no stroke,
// removes existing content, sets text properties, and adds a paragraph.

var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 80 * 36000, 50 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Remove all existing elements in the shape's content
oDocContent.RemoveAllElements();

// Create text properties
var oTextPr = Api.CreateTextPr();
oTextPr.SetFontSize(30);
oTextPr.SetBold(true);

// Create a paragraph and set its alignment to left
var oParagraph = Api.CreateParagraph();
oParagraph.SetJc("left");

// Add text to the paragraph
oParagraph.AddText("This is a sample text with the font size set to 30 and the font weight set to bold.");

// Apply text properties to the paragraph
oParagraph.SetTextPr(oTextPr);

// Add the paragraph to the shape's content
oDocContent.Push(oParagraph);
```