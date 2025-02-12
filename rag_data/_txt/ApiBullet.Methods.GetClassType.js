# Insert a bulleted shape into the worksheet and display its class type  
# Вставить маркированную фигуру в лист и отобразить её тип класса

## VBA Code

```vba
' This VBA code inserts a flowchart shape into the active worksheet,
' applies a solid fill and no stroke, adds bulleted paragraphs, and displays the class type.

Sub InsertBulletedShape()
    Dim oWorksheet As Worksheet
    Dim oShape As Shape
    Dim oFill As FillFormat
    Dim oStroke As LineFormat
    Dim oTextFrame As TextFrame2
    Dim oTextRange As TextRange2
    Dim oParagraph As ParagraphFormat
    Dim sClassType As String
    
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Add a shape to the worksheet
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, _
        120, 35, 200, 100) ' Position and size in points
    
    ' Apply solid fill with RGB color (255, 111, 61)
    Set oFill = oShape.Fill
    oFill.Visible = msoTrue
    oFill.Solid
    oFill.ForeColor.RGB = RGB(255, 111, 61)
    
    ' Apply no stroke
    Set oStroke = oShape.Line
    oStroke.Visible = msoFalse
    
    ' Access the text frame of the shape
    Set oTextFrame = oShape.TextFrame2
    oTextFrame.TextRange.Text = ""
    
    ' Add first paragraph with bullet
    Set oTextRange = oTextFrame.TextRange
    oTextRange.Paragraphs.Add
    With oTextRange.Paragraphs(1).ParagraphFormat
        .Bullet.Visible = msoTrue
        .Bullet.Character = 49 ' ASCII for '1'
        .Bullet.Type = msoBulletNumbered
        .Bullet.Style = msoBulletArabicPeriod
    End With
    oTextRange.Paragraphs(1).Text = "This is an example of the bulleted paragraph."
    
    ' Add second paragraph with bullet
    oTextRange.Paragraphs.Add
    With oTextRange.Paragraphs(2).ParagraphFormat
        .Bullet.Visible = msoTrue
        .Bullet.Character = 49
        .Bullet.Type = msoBulletNumbered
        .Bullet.Style = msoBulletArabicPeriod
    End With
    oTextRange.Paragraphs(2).Text = "This is an example of the bulleted paragraph."
    
    ' Get the class type (for demonstration, using the shape type)
    sClassType = TypeName(oShape)
    
    ' Add third paragraph with class type
    oTextRange.Paragraphs.Add
    With oTextRange.Paragraphs(3).ParagraphFormat
        .Alignment = msoAlignLeft
    End With
    oTextRange.Paragraphs(3).Text = "Class Type = " & sClassType
End Sub
```

## JavaScript Code

```javascript
// This example shows how to get a type of the ApiBullet class and insert it into the table.
// Этот пример демонстрирует, как получить тип класса ApiBullet и вставить его в таблицу.

var oWorksheet = Api.GetActiveSheet();
// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
// Add a shape to the worksheet with specified properties
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
// Get the content of the shape
var oDocContent = oShape.GetContent();
// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);
// Create a numbering bullet of type "ArabicParenR"
var oBullet = Api.CreateNumbering("ArabicParenR", 1);
// Set the bullet to the paragraph
oParagraph.SetBullet(oBullet);
// Add text to the paragraph
oParagraph.AddText(" This is an example of the bulleted paragraph.");
// Create a new paragraph and set the bullet
oParagraph = Api.CreateParagraph();
oParagraph.SetBullet(oBullet);
// Add text to the new paragraph
oParagraph.AddText(" This is an example of the bulleted paragraph.");
// Push the new paragraph to the document content
oDocContent.Push(oParagraph);
// Get the class type of the bullet
var sClassType = oBullet.GetClassType();
// Create another paragraph, set alignment to left, and add text with class type
oParagraph = Api.CreateParagraph();
oParagraph.SetJc("left");
oParagraph.AddText("Class Type = " + sClassType);
// Push the paragraph to the document content
oDocContent.Push(oParagraph);
```