# Description

**English:**
This code adds a shape to the active worksheet, inserts two paragraphs of text into the shape's content, and sets the left indentation of the first paragraph to 2 inches.

**Russian:**
Этот код добавляет форму на активный лист, вставляет два абзаца текста в содержимое формы и устанавливает отступ слева первого абзаца на 2 дюйма.

```vba
' VBA Code
' This code adds a shape to the active worksheet, inserts two paragraphs of text,
' and sets the left indentation of the first paragraph.

Sub AddShapeWithIndentedParagraph()
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet

    ' Add a flowchart storage shape to the worksheet with specific position and size
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(Type:=msoShapeFlowchartStorage, _
        Left:=120, Top:=70, Width:=200, Height:=100) ' Adjust size as needed

    ' Set the fill color to RGB(255, 111, 61)
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61)
        .Solid
    End With

    ' Set the line to no visible fill and width 0
    With oShape.Line
        .Visible = msoFalse
        .Weight = 0
    End With

    ' Add text to the shape
    With oShape.TextFrame2
        .TextRange.Text = "This is a paragraph with the indent of 2 inches set to it. " & _
                          "These sentences are used to add lines for demonstrative purposes." & vbCrLf & _
                          "This is a paragraph without any indent set to it. " & _
                          "These sentences are used to add lines for demonstrative purposes."

        ' Set left indentation of the first paragraph to 2 inches (144 points)
        .TextRange.Paragraphs(1).ParagraphFormat.LeftIndent = 144

        ' The second paragraph remains without indentation
    End With
End Sub
```

```javascript
// OnlyOffice JavaScript Code
// This code adds a shape to the active worksheet, inserts two paragraphs of text,
// and sets the left indentation of the first paragraph to 2 inches.

var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with width 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a flowchart storage shape to the worksheet with specified position and size
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 
    120 * 36000, // Left position
    70 * 36000,  // Top position
    200 * 36000, // Width
    100 * 36000, // Height
    oFill, 
    oStroke, 
    0, 
    0, 
    0, 
    0);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph from the shape's content
var oParagraph = oDocContent.GetElement(0);

// Add text to the first paragraph
oParagraph.AddText("This is a paragraph with the indent of 2 inches set to it. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");

// Set left indentation of the first paragraph to 2880 twips (2 inches)
oParagraph.SetIndLeft(2880);

// Create a new paragraph
oParagraph = Api.CreateParagraph();

// Add text to the second paragraph
oParagraph.AddText("This is a paragraph without any indent set to it. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");

// Push the second paragraph to the document content
oDocContent.Push(oParagraph);
```