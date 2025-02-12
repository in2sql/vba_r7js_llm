# Code Description / Описание кода

**English:**  
This code retrieves the active worksheet, creates a solid fill and stroke, adds a shape to the sheet, and inserts bulleted paragraphs into the shape's content. It also retrieves the class type of the bullet and displays it.

**Russian:**  
Этот код получает активный лист, создает заливку и обводку, добавляет фигуру на лист и вставляет маркированные абзацы в содержимое фигуры. Также он получает тип класса маркера и отображает его.

```javascript
// JavaScript code using OnlyOffice API

// This example shows how to get a type of the ApiBullet class and insert it into the table.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill with RGB color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
var oShape = oWorksheet.AddShape(
    "flowChartOnlineStorage", 
    120 * 36000, 
    35 * 36000, 
    oFill, 
    oStroke, 
    0, 
    2 * 36000, 
    0, 
    3 * 36000
); // Add a shape to the worksheet
var oDocContent = oShape.GetContent(); // Get the content of the shape
var oParagraph = oDocContent.GetElement(0); // Get the first paragraph element
var oBullet = Api.CreateNumbering("ArabicParenR", 1); // Create a bullet numbering style
oParagraph.SetBullet(oBullet); // Apply the bullet to the paragraph
oParagraph.AddText(" This is an example of the bulleted paragraph."); // Add text to the paragraph
oParagraph = Api.CreateParagraph(); // Create a new paragraph
oParagraph.SetBullet(oBullet); // Apply the same bullet to the new paragraph
oParagraph.AddText(" This is an example of the bulleted paragraph."); // Add text to the new paragraph
oDocContent.Push(oParagraph); // Add the new paragraph to the document content
var sClassType = oBullet.GetClassType(); // Get the class type of the bullet
oParagraph = Api.CreateParagraph(); // Create another new paragraph
oParagraph.SetJc("left"); // Set paragraph alignment to left
oParagraph.AddText("Class Type = " + sClassType); // Add text displaying the class type
oDocContent.Push(oParagraph); // Add this paragraph to the document content
```

```vba
' VBA code equivalent for Excel

' This VBA subroutine adds a shape to the active worksheet, fills it with a specific color,
' adds bulleted text, and displays the type of bullet used.

Sub AddShapeWithBullets()
    ' Get the active worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Create and add a shape (Flowchart Online Storage)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartOnlineStorage, 120, 35, 200, 150) ' Position (Left, Top) in points and size (Width, Height)

    ' Set the fill color to RGB(255, 111, 61)
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61)
        .Solid
    End With

    ' Set the stroke to no fill with weight 0
    With shp.Line
        .Visible = msoTrue
        .Weight = 0 ' No line weight
        .ForeColor.RGB = RGB(255, 255, 255) ' White color (assuming no visible stroke)
    End With

    ' Add bulleted text to the shape
    With shp.TextFrame2.TextRange
        .Text = "This is an example of the bulleted paragraph." & vbCrLf & _
                "This is an example of the bulleted paragraph."
        ' Apply bullets to each paragraph
        .ParagraphFormat.Bullet.Visible = msoTrue
        .ParagraphFormat.Bullet.Character = 8226 ' Unicode for bullet character
    End With

    ' Display the bullet type (using Unicode bullet as VBA does not have a direct class type equivalent)
    Dim bulletType As String
    bulletType = "Class Type = Bullet"
    With shp.TextFrame2.TextRange
        .Text = .Text & vbCrLf & bulletType
        ' Align the last paragraph to left
        .Paragraphs(.Paragraphs.Count).ParagraphFormat.Alignment = msoAlignLeft
    End With
End Sub
```