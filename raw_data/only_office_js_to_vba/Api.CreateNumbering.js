### Description / Описание

**English:**  
This code creates a shape on the active worksheet, sets its fill and stroke properties, and adds two numbered paragraphs with Arabic parenthesis numbering.

**Russian:**  
Этот код создает фигуру на активном листе, устанавливает свойства заливки и контура, а также добавляет два пронумерованных абзаца с арабской нумерацией в скобках.

```vba
' VBA Code to create a shape with numbered paragraphs

Sub CreateNumberedShape()
    Dim oWorksheet As Worksheet
    Dim oShape As Shape
    Dim oTextFrame As TextFrame2
    Dim oTextRange As TextRange2
    Dim oParagraph As TextRange2
    Dim bulletStyle As String
    
    ' Set the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Add a shape to the worksheet
    Set oShape = oWorksheet.Shapes.AddShape( _
        Type:=msoShapeFlowChartOnlineStorage, _
        Left:=120, _
        Top:=35, _
        Width:=200, _
        Height:=100)
    
    ' Set the fill color (RGB: 255, 111, 61)
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61)
        .Solid
    End With
    
    ' Set the line properties (no fill)
    With oShape.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0) ' No fill equivalent
        .Transparency = 1 ' Fully transparent
    End With
    
    ' Access the text frame of the shape
    Set oTextFrame = oShape.TextFrame2
    oTextFrame.TextRange.Text = ""
    
    ' Define bullet style
    bulletStyle = "1)"
    
    ' Add first paragraph with bullet
    Set oParagraph = oTextFrame.TextRange.Paragraphs.Add
    With oParagraph
        .Text = " This is an example of the numbered paragraph."
        .ParagraphFormat.Bullet.Visible = msoTrue
        .ParagraphFormat.Bullet.Character = Asc("1") ' Arabic numbering
    End With
    
    ' Add second paragraph with bullet
    Set oParagraph = oTextFrame.TextRange.Paragraphs.Add
    With oParagraph
        .Text = " This is an example of the numbered paragraph."
        .ParagraphFormat.Bullet.Visible = msoTrue
        .ParagraphFormat.Bullet.Character = Asc("1") ' Arabic numbering
    End With
End Sub
```

```javascript
// JavaScript Code to create a shape with numbered paragraphs using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet
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
);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph
var oParagraph = oDocContent.GetElement(0);

// Create a numbering bullet with Arabic parentheses
var oBullet = Api.CreateNumbering("ArabicParenR", 1);

// Set the bullet for the first paragraph
oParagraph.SetBullet(oBullet);

// Add text to the first paragraph
oParagraph.AddText(" This is an example of the numbered paragraph.");

// Create a new paragraph
oParagraph = Api.CreateParagraph();

// Set the bullet for the new paragraph
oParagraph.SetBullet(oBullet);

// Add text to the new paragraph
oParagraph.AddText(" This is an example of the numbered paragraph.");

// Add the new paragraph to the document content
oDocContent.Push(oParagraph);
```