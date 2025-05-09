**Description / Описание**

English: This code adds a flowchart shape to the active worksheet, sets its fill and stroke properties, adds a paragraph with a left indentation of 2 inches, retrieves the indentation value, and displays it in a new paragraph.

Russian: Этот код добавляет фигуру блок-схемы на активный рабочий лист, устанавливает свойства заливки и контура, добавляет абзац с отступом слева в 2 дюйма, получает значение отступа и отображает его в новом абзаце.

```vba
' VBA Code to add a flowchart shape, set properties, and manage paragraph indentations

Sub AddFlowchartWithIndent()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim txtFrame As TextFrame2
    Dim para As TextRange2
    Dim indLeft As Single
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Add a flowchart shape to the worksheet
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartInternalStorage, _
        120 * 360, 70 * 360, 200, 100) ' Positions and size in points
    
    ' Set the fill color to RGB(255, 111, 61)
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61)
        .Solid
    End With
    
    ' Remove the line (stroke)
    shp.Line.Visible = msoFalse
    
    ' Access the text frame of the shape
    Set txtFrame = shp.TextFrame2
    
    ' Add text with left indentation of 2 inches (144 points)
    With txtFrame.TextRange
        .Text = "This is a paragraph with the indent of 2 inches set to it. " & _
                "These sentences are used to add lines for demonstrative purposes. "
        .ParagraphFormat.LeftIndent = 144 ' 2 inches in points
    End With
    
    ' Retrieve the left indentation
    indLeft = txtFrame.TextRange.ParagraphFormat.LeftIndent
    
    ' Add a new paragraph to display the indentation
    With txtFrame.TextRange
        .InsertAfter vbCrLf & "Left indent: " & indLeft & " points"
    End With
End Sub
```

```javascript
// JavaScript Code using OnlyOffice API to add a flowchart shape, set properties, and manage paragraph indentations

// This example shows how to get the paragraph left side indentation.
// English: Adds a flowchart shape, sets its fill and stroke, adds a paragraph with indentation, and displays the indentation value.
// Russian: Добавляет фигуру блок-схемы, устанавливает заливку и контур, добавляет абзац с отступом и отображает значение отступа.

var oWorksheet = Api.GetActiveSheet();
// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
// Add the shape to the worksheet with specified parameters
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
// Get the content of the shape
var oDocContent = oShape.GetContent();
// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);
// Add text to the paragraph
oParagraph.AddText("This is a paragraph with the indent of 2 inches set to it. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
// Set the left indentation to 2880 (units depend on API)
oParagraph.SetIndLeft(2880);
// Get the left indentation value
var nIndLeft = oParagraph.GetIndLeft();
// Create a new paragraph
oParagraph = Api.CreateParagraph();
// Add text displaying the left indentation
oParagraph.AddText("Left indent: " + nIndLeft);
// Push the new paragraph to the document content
oDocContent.Push(oParagraph); 
```