# Description / Описание

**English:**  
This code adds a shape to the active worksheet, sets its fill color and stroke, and adds paragraphs of text with specified indentation.

**Русский:**  
Этот код добавляет фигуру на активный лист, устанавливает её цвет заливки и обводку, а также добавляет абзацы текста с указанным отступом.

```vba
' VBA Code to add a shape with specific fill and stroke, and add text with indentation

Sub AddShapeWithIndentedText()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim txtBox As Shape
    Dim txtRange As TextRange
    Dim para1 As TextRange
    Dim para2 As TextRange
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Add a rectangle shape to the worksheet
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartDatabase, 120, 70, 300, 150)
    
    ' Set the fill color to RGB (255, 111, 61)
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61)
        .Solid
    End With
    
    ' Set the line (stroke) to no fill
    With shp.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 255) ' Assuming no color for no fill
        .Weight = 0
    End With
    
    ' Add text to the shape
    shp.TextFrame2.TextRange.Text = ""
    
    ' Add first paragraph with indentation
    Set para1 = shp.TextFrame2.TextRange.Paragraphs(1)
    para1.Text = "This is a paragraph with the indent of 2 inches set to it. " & _
                "These sentences are used to add lines for demonstrative purposes."
    para1.IndentLevel = 2 ' Setting indentation (approx 2 inches)
    
    ' Add second paragraph without indentation
    Set para2 = shp.TextFrame2.TextRange.Paragraphs.Add
    para2.Text = "This is a paragraph without any indent set to it. " & _
                "These sentences are used to add lines for demonstrative purposes."
    para2.IndentLevel = 0
End Sub
```

```javascript
// JavaScript Code to add a shape with specific fill and stroke, and add text with indentation

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill color with RGB (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with width 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified type and dimensions
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph in the shape's content
var oParagraph = oDocContent.GetElement(0);

// Add text with indentation to the first paragraph
oParagraph.AddText("This is a paragraph with the indent of 2 inches set to it. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.SetIndLeft(2880); // Set left indentation

// Create a new paragraph without indentation
oParagraph = Api.CreateParagraph();
oParagraph.AddText("This is a paragraph without any indent set to it. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");

// Add the new paragraph to the shape's content
oDocContent.Push(oParagraph);
```