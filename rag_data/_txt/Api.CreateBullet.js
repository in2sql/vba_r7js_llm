## Description

**English:** This code creates a bullet for a paragraph using OnlyOffice API and provides the equivalent VBA code in Excel.

**Russian:** Этот код создает маркер для абзаца с использованием OnlyOffice API и предоставляет эквивалентный код VBA в Excel.

```javascript
// This example creates a bullet for a paragraph.
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();
// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
// Create a stroke with width 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
// Add a shape to the worksheet with specified parameters
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
// Get the content of the shape
var oDocContent = oShape.GetContent();
// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);
// Create a bullet with "-"
var oBullet = Api.CreateBullet("-");
// Set the bullet to the paragraph
oParagraph.SetBullet(oBullet);
// Add text to the paragraph
oParagraph.AddText(" This is an example of the bulleted paragraph.");
```

```vba
' This example creates a bullet for a paragraph in Excel using VBA.
Sub CreateBulletedParagraph()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Add a shape to the worksheet
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartStorage, 120, 35, 200, 100) ' Width and height are in points
    
    ' Set the fill color (RGB: 255, 111, 61)
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61)
    End With
    
    ' Set the line (stroke) to no line
    With oShape.Line
        .Visible = msoFalse
    End With
    
    ' Add text to the shape
    With oShape.TextFrame2.TextRange
        ' Enable bullet formatting
        .ParagraphFormat.Bullet.Visible = msoTrue
        ' Set the bullet character to "-"
        .ParagraphFormat.Bullet.Character = Asc("-")
        ' Add the paragraph text
        .Text = "This is an example of the bulleted paragraph."
    End With
End Sub
```