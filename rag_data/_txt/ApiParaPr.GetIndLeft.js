```plaintext
// This code demonstrates how to add a shape to the active worksheet, set the left indentation of a paragraph within the shape, add text to the paragraph, retrieve the indentation value, create a new paragraph displaying the indentation, and add it to the shape.
// Этот код демонстрирует, как добавить фигуру на активный лист, установить отступ слева для абзаца внутри фигуры, добавить текст к абзацу, получить значение отступа, создать новый абзац с отображением отступа и добавить его к фигуре.

' VBA Code Equivalent
Sub AddShapeAndSetIndent()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim txtFrame As TextFrame
    Dim txtRange As TextRange
    Dim para As ParagraphFormat
    Dim indentValue As Single
    
    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Add a shape to the worksheet
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartOnlineStorage, 120, 70, 300, 200)
    
    ' Get the text frame of the shape
    Set txtFrame = shp.TextFrame
    
    ' Set the text of the shape
    txtFrame.Characters.Text = "This is the first paragraph with the indent of 2 inches set to it. " & _
                               "This indent is set by the paragraph style. No paragraph inline style is applied."
    
    ' Get the paragraph format
    Set para = txtFrame.MarginLeft
    Set para = txtFrame.Characters(1, 100).ParagraphFormat
    
    ' Set left indentation to 2 inches (144 points)
    para.LeftIndent = 144
    
    ' Retrieve the left indentation value
    indentValue = para.LeftIndent
    
    ' Add a new paragraph with the indentation value
    txtFrame.Characters.Text = txtFrame.Characters.Text & vbCrLf & "Left indent: " & indentValue
End Sub
```

```javascript
// This code demonstrates how to add a shape to the active worksheet, set the left indentation of a paragraph within the shape, add text to the paragraph, retrieve the indentation value, create a new paragraph displaying the indentation, and add it to the shape.
// Этот код демонстрирует, как добавить фигуру на активный лист, установить отступ слева для абзаца внутри фигуры, добавить текст к абзацу, получить значение отступа, создать новый абзац с отображением отступа и добавить его к фигуре.

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph in the shape's content
var oParagraph = oDocContent.GetElement(0);

// Get the paragraph properties
var oParaPr = oParagraph.GetParaPr();

// Set the left indentation to 2880 (assuming the unit)
oParaPr.SetIndLeft(2880);

// Add text to the paragraph
oParagraph.AddText("This is the first paragraph with the indent of 2 inches set to it. ");
oParagraph.AddText("This indent is set by the paragraph style. No paragraph inline style is applied. ");

// Retrieve the left indentation value
var nIndLeft = oParaPr.GetIndLeft();

// Create a new paragraph
oParagraph = Api.CreateParagraph();

// Add text displaying the indentation value
oParagraph.AddText("Left indent: " + nIndLeft);

// Add the new paragraph to the shape's content
oDocContent.Push(oParagraph);
```