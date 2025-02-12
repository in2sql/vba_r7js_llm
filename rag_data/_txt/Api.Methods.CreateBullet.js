**Description / Описание:**  
This script creates a bullet point for a paragraph by adding a shape to the active worksheet, setting its fill and stroke properties, and inserting a bulleted paragraph with specified text.  
Этот скрипт создает маркированный пункт для абзаца, добавляя фигуру на активный лист, устанавливая свойства заливки и обводки, и вставляя маркированный абзац с заданным текстом.

```vba
' VBA Code to create a bullet for a paragraph in Excel

Sub CreateBulletedParagraph()
    Dim oWorksheet As Worksheet
    Dim oShape As Shape
    Dim oParagraph As TextRange
    Dim oBullet As String
    
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Add a shape to the worksheet
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, _
        120, 35, 200, 100) ' (Left, Top, Width, Height)
    
    ' Get the text frame of the shape
    With oShape.TextFrame
        .Characters.Text = "This is an example of the bulleted paragraph."
        Set oParagraph = .Characters.Paragraphs(1)
        oBullet = "• " ' Define bullet symbol
        ' Set the paragraph to be bulleted
        oParagraph.Text = oBullet & " This is an example of the bulleted paragraph."
    End With
End Sub
```

```javascript
// JavaScript Code to create a bullet for a paragraph using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with width 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified properties
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Create a bullet with "-" as the symbol
var oBullet = Api.CreateBullet("-");

// Set the bullet for the paragraph
oParagraph.SetBullet(oBullet);

// Add text to the paragraph
oParagraph.AddText(" This is an example of the bulleted paragraph."); 
```