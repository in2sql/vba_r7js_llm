**Description / Описание**

This code demonstrates how to create a shape in the active worksheet, set the first line indentation of a paragraph within the shape, add multiple lines of text, and display the indentation value.

Этот код демонстрирует, как создать фигуру в активном листе, установить отступ первой строки абзаца внутри фигуры, добавить несколько строк текста и отобразить значение отступа.

```vba
' VBA Code to create a shape, set paragraph indentation, and add text in Excel

Sub CreateShapeWithIndentedParagraph()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Create a solid fill color (RGB: 255, 111, 61)
    Dim fillColor As Long
    fillColor = RGB(255, 111, 61)
    
    ' Add a rectangle shape to the worksheet
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, 120, 70, 360, 70)
    shp.Fill.ForeColor.RGB = fillColor
    shp.Line.Visible = msoFalse ' No stroke
    
    ' Add a text box within the shape
    shp.TextFrame2.TextRange.ParagraphFormat.FirstLineIndent = 36 ' 1 inch in points
    shp.TextFrame2.TextRange.Text = "This is the first paragraph with the indent of 1 inch set to the first line. " & _
                                    "This indent is set by the paragraph style. No paragraph inline style is applied. " & _
                                    "These sentences are used to add lines for demonstrative purposes. " & _
                                    "These sentences are used to add lines for demonstrative purposes. " & _
                                    "These sentences are used to add lines for demonstrative purposes."
    
    ' Retrieve the first line indent value
    Dim indentValue As Single
    indentValue = shp.TextFrame2.TextRange.ParagraphFormat.FirstLineIndent
    
    ' Add another shape to display the indent value
    Dim shpIndent As Shape
    Set shpIndent = ws.Shapes.AddShape(msoShapeRectangle, 120, 150, 360, 30)
    shpIndent.Fill.Visible = msoFalse
    shpIndent.Line.Visible = msoFalse
    shpIndent.TextFrame2.TextRange.Text = "First line indent: " & indentValue & " points"
End Sub
```

```javascript
// JavaScript Code to create a shape, set paragraph indentation, and add text in OnlyOffice

// This example shows how to create a shape, set the paragraph's first line indentation, add text, and display the indentation value.
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill color (RGB: 255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the document content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph
var oParagraph = oDocContent.GetElement(0);

// Get the paragraph properties
var oParaPr = oParagraph.GetParaPr();

// Set the first line indentation to 1440 (twips)
oParaPr.SetIndFirstLine(1440);

// Add multiple lines of text to the paragraph
oParagraph.AddText("This is the first paragraph with the indent of 1 inch set to the first line. ");
oParagraph.AddText("This indent is set by the paragraph style. No paragraph inline style is applied. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.");

// Retrieve the first line indent value
var nIndFirstLine = oParaPr.GetIndFirstLine();

// Create a new paragraph to display the indent value
oParagraph = Api.CreateParagraph();
oParagraph.AddText("First line indent: " + nIndFirstLine);

// Push the new paragraph to the document content
oDocContent.Push(oParagraph);
```