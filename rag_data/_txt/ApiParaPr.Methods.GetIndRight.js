# Example of Setting Right Indentation for a Paragraph / Пример установки правого отступа для абзаца

This code adds a shape to the active worksheet, sets the right indentation and alignment of the first paragraph within the shape, adds text to it, retrieves the right indentation value, and appends a new paragraph displaying this value.

Этот код добавляет фигуру на активный рабочий лист, устанавливает правый отступ и выравнивание первого абзаца внутри фигуры, добавляет к нему текст, получает значение правого отступа и добавляет новый абзац, отображающий это значение.

```javascript
// This example shows how to set the paragraph right side indentation.
// Этот пример показывает, как установить правый отступ абзаца.

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill with RGB color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
// Add a shape to the worksheet with specified parameters
var oShape = oWorksheet.AddShape(
    "flowChartOnlineStorage",
    120 * 36000,
    70 * 36000,
    oFill,
    oStroke,
    0,
    2 * 36000,
    0,
    3 * 36000
);
var oDocContent = oShape.GetContent(); // Get the content of the shape
var oParagraph = oDocContent.GetElement(0); // Get the first paragraph
var oParaPr = oParagraph.GetParaPr(); // Get paragraph properties
oParaPr.SetIndRight(2880); // Set right indentation to 2880 (e.g., twips)
oParaPr.SetJc("right"); // Set justification to right
// Add text to the paragraph
oParagraph.AddText("This is the first paragraph with the right offset of 2 inches set to it. ");
oParagraph.AddText("This indent is set by the paragraph style. No paragraph inline style is applied. ");
var nIndRight = oParaPr.GetIndRight(); // Get the right indentation value
oParagraph = Api.CreateParagraph(); // Create a new paragraph
oParagraph.AddText("Right indent: " + nIndRight); // Add text displaying the right indentation
oDocContent.Push(oParagraph); // Append the new paragraph to the document content
```

```vba
' This example shows how to set the paragraph right side indentation.
' Этот пример показывает, как установить правый отступ абзаца.

Sub SetRightIndentation()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' Get the active worksheet
    
    ' Add a rectangle shape to the worksheet
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartStorage, 120, 70, 200, 100)
    
    ' Set the fill color to RGB(255, 111, 61)
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61)
        .Solid
    End With
    
    ' Remove the line (stroke) from the shape
    With shp.Line
        .Visible = msoFalse
    End With
    
    ' Access the text frame of the shape
    With shp.TextFrame2
        .TextRange.ParagraphFormat.Alignment = msoAlignRight ' Set alignment to right
        .TextRange.ParagraphFormat.RightIndent = 28.8 ' Set right indentation (points)
        
        ' Add text to the shape
        .TextRange.Text = "This is the first paragraph with the right offset of 2 inches set to it." & vbCrLf & _
                          "This indent is set by the paragraph style. No paragraph inline style is applied."
                          
        ' Retrieve the right indentation value
        Dim rightIndent As Single
        rightIndent = .TextRange.ParagraphFormat.RightIndent
        
        ' Add a new paragraph displaying the right indentation
        .TextRange.Text = .TextRange.Text & vbCrLf & "Right indent: " & rightIndent & " points"
    End With
End Sub
```