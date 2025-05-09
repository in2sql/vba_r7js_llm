---

**Description / Описание**

English:  
This script adds a shape to the active worksheet, sets its fill and stroke properties, and adds a bulleted paragraph with specified text.

Russian:  
Этот скрипт добавляет фигуру на активный лист, устанавливает свойства заливки и обводки, а также добавляет абзац с маркером и указанным текстом.

---

#### VBA Code / Код VBA

```vba
' This VBA script adds a shape to the active worksheet, sets its fill and stroke properties,
' and adds a bulleted paragraph with specific text.

Sub AddBulletedParagraph()
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Create RGB color
    Dim fillColor As Long
    fillColor = RGB(255, 111, 61)
    
    ' Add shape to worksheet
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartDatabase, _
        120, 35, 200, 150) ' Width and Height are in points
    
    ' Set shape fill
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = fillColor
        .Solid
    End With
    
    ' Set shape line (stroke)
    With oShape.Line
        .Visible = msoTrue
        .Weight = 0
        .ForeColor.RGB = RGB(255, 255, 255) ' No fill (white)
    End With
    
    ' Add text with bullet
    With oShape.TextFrame2.TextRange
        .ParagraphFormat.Bullet.Visible = msoTrue
        .ParagraphFormat.Bullet.Character = 8226 ' Unicode bullet character
        .Text = " This is an example of the bulleted paragraph."
    End With
End Sub
```

#### OnlyOffice JS Code / Код OnlyOffice JS

```javascript
// This JavaScript code sets bullet points for a paragraph by creating a shape in the active sheet,
// applying a solid fill and stroke, and adding a bulleted paragraph with specific text.

var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified properties
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Create a bullet with "-" character
var oBullet = Api.CreateBullet("-");

// Set the bullet to the paragraph
oParagraph.SetBullet(oBullet);

// Add text to the paragraph
oParagraph.AddText(" This is an example of the bulleted paragraph.");
```