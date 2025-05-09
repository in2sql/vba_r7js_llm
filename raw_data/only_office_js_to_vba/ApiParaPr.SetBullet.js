### Description / Описание
**English:**  
This code adds a shape to the active worksheet in OnlyOffice, sets its fill and stroke properties, and adds a bulleted paragraph with specified text.

**Русский:**  
Этот код добавляет фигуру на активный рабочий лист в OnlyOffice, устанавливает свойства заливки и обводки, а также добавляет маркированный абзац с указанным текстом.

```vba
' VBA Code to add a shape with a bulleted paragraph in Excel

Sub AddBulletedShape()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Add a shape to the worksheet
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartDecision, _
        120, 35, 200, 100) ' Adjust size and position as needed
    
    ' Set the fill color (RGB: 255, 111, 61)
    oShape.Fill.ForeColor.RGB = RGB(255, 111, 61)
    
    ' Remove the outline
    oShape.Line.Visible = msoFalse
    
    ' Add text to the shape
    oShape.TextFrame.Characters.Text = "This is an example of the bulleted paragraph."
    
    ' Set bullet style
    With oShape.TextFrame2.TextRange.ParagraphFormat
        .Bullet.Visible = msoTrue
        .Bullet.Character = 45 ' ASCII code for "-"
    End With
End Sub
```

```javascript
// JavaScript Code to add a shape with a bulleted paragraph in OnlyOffice

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet (flowChartOnlineStorage type)
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Get the paragraph properties
var oParaPr = oParagraph.GetParaPr();

// Create a bullet with "-"
var oBullet = Api.CreateBullet("-");

// Set the bullet to the paragraph
oParaPr.SetBullet(oBullet);

// Add text to the paragraph
oParagraph.AddText(" This is an example of the bulleted paragraph.");
```