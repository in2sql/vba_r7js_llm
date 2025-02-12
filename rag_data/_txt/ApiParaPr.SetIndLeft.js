# Description / Описание

**English:**  
This code adds a shape to the active worksheet, sets its fill and stroke properties, adjusts the left indentation of the first paragraph within the shape, and adds multiple lines of text to demonstrate the indentation.

**Russian:**  
Этот код добавляет фигуру на активный лист, устанавливает свойства заливки и обводки, корректирует левый отступ первого абзаца внутри фигуры и добавляет несколько строк текста для демонстрации отступа.

```javascript
// Add a shape to the active worksheet with specified fill and stroke
var oWorksheet = Api.GetActiveSheet();
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill with RGB color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape and adjust paragraph indentation
var oDocContent = oShape.GetContent();
var oParagraph = oDocContent.GetElement(0);
var oParaPr = oParagraph.GetParaPr();
oParaPr.SetIndLeft(2880); // Set left indentation to 2 inches

// Add multiple lines of text to the paragraph
oParagraph.AddText("This is the first paragraph with the indent of 2 inches set to it. ");
oParagraph.AddText("This indent is set by the paragraph style. No paragraph inline style is applied. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
```

```vba
' Add a shape to the active worksheet with specified fill and line
Sub AddShapeWithIndent()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Add a shape of type FlowChartOnlineStorage with specified dimensions
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartOnlineStorage, 120, 70, 200, 150)
    
    ' Set the fill color to RGB(255, 111, 61)
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61)
        .Solid
    End With
    
    ' Remove the line (stroke)
    With shp.Line
        .Visible = msoFalse
    End With
    
    ' Add text to the shape
    With shp.TextFrame2.TextRange
        ' Set left indentation to 2 inches (144 points)
        .ParagraphFormat.LeftIndent = 144
        ' Add multiple lines of text
        .Text = "This is the first paragraph with the indent of 2 inches set to it. " & vbCrLf & _
                "This indent is set by the paragraph style. No paragraph inline style is applied. " & vbCrLf & _
                "These sentences are used to add lines for demonstrative purposes."
    End With
End Sub
```