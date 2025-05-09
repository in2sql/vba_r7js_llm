**Description / Описание**
This code adds a flowchart shape to the active sheet, inserts paragraphs with and without first-line indentation, and sets the indentation for the first paragraph.

Этот код добавляет фигуру блок-схемы на активный лист, вставляет абзацы с отступом первой строки и без него, а также устанавливает отступ для первого абзаца.

```vba
' VBA Code to add a shape and manipulate paragraphs with indentation

Sub AddFlowChartShape()
    Dim oWorksheet As Worksheet
    Dim oShape As Shape
    Dim oDocContent As Object
    Dim oParagraph As Object
    
    ' Set the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Add a flowchart shape to the worksheet
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, _
        120 * 72, 70 * 72, 200, 150) ' Width and Height are in points (72 points = 1 inch)
    
    ' Access the text frame of the shape
    Set oDocContent = oShape.TextFrame2.TextRange
    
    ' Add text with first line indentation
    oDocContent.Text = "This is a paragraph with the indent of 1 inch set to the first line. " & _
                      "These sentences are used to add lines for demonstrative purposes. " & _
                      "These sentences are used to add lines for demonstrative purposes. " & _
                      "These sentences are used to add lines for demonstrative purposes."
    
    ' Set first line indent to 1 inch (72 points)
    oDocContent.ParagraphFormat.FirstLineIndent = 72
    
    ' Add a new paragraph without indentation
    oDocContent.Text = oDocContent.Text & vbCrLf & _
                      "This is a paragraph without any indent set to the first line. " & _
                      "These sentences are used to add lines for demonstrative purposes. " & _
                      "These sentences are used to add lines for demonstrative purposes."
    ' The new paragraph inherits default indentation (no indent)
End Sub
```

```javascript
// JavaScript Code using OnlyOffice API to add a shape and manipulate paragraphs with indentation

// Description: Adds a flowchart shape, inserts paragraphs with and without first-line indentation, and sets the indentation for the first paragraph.

// This example sets the paragraph first line indentation.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill with RGB color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000); // Add the shape with specified parameters
var oDocContent = oShape.GetContent(); // Get the content of the shape

// Get the first paragraph and add text with indentation
var oParagraph = oDocContent.GetElement(0);
oParagraph.AddText("This is a paragraph with the indent of 1 inch set to the first line. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.");
oParagraph.SetIndFirstLine(1440); // Set first line indent to 1 inch (1440 twips)

// Create a new paragraph without indentation and add text
oParagraph = Api.CreateParagraph();
oParagraph.AddText("This is a paragraph without any indent set to the first line. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.");
oDocContent.Push(oParagraph); // Add the new paragraph to the document content
```