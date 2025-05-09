### Description / Описание

**English:**  
This code adds a shape to the active worksheet with a specified fill color and no stroke. It inserts a paragraph of text with a first-line indentation of 1 inch and appends a line displaying the indentation value.

**Русский:**  
Этот код добавляет фигуру на активный лист с заданным цветом заливки и без обводки. Он вставляет абзац текста с отступом первой строки в 1 дюйм и добавляет строку, отображающую значение этого отступа.

```vba
' VBA Code
Sub AddShapeWithIndentedParagraph()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim txtRange As TextRange2
    Dim para As ParagraphFormat2
    Dim indentValue As Single
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Add a rectangle shape with specified dimensions
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, 120, 70, 200, 100)
    
    ' Set fill color (RGB: 255, 111, 61)
    shp.Fill.ForeColor.RGB = RGB(255, 111, 61)
    
    ' Remove the stroke
    shp.Line.Visible = msoFalse
    
    ' Access the text frame
    With shp.TextFrame2
        .TextRange.Text = "This is a paragraph with the indent of 1 inch set to the first line. " & _
                         "These sentences are used to add lines for demonstrative purposes. " & _
                         "These sentences are used to add lines for demonstrative purposes. " & _
                         "These sentences are used to add lines for demonstrative purposes."
        
        ' Set first line indent (1 inch = 72 points)
        .TextRange.ParagraphFormat.FirstLineIndent = 72
        
        ' Retrieve the first line indent value
        indentValue = .TextRange.ParagraphFormat.FirstLineIndent
        
        ' Append a new paragraph with indentation info
        .TextRange.Text = .TextRange.Text & vbCrLf & "First line indent: " & indentValue & " points"
    End With
End Sub
```

```javascript
// JavaScript Code
// This example shows how to get the paragraph first line indentation.
var oWorksheet = Api.GetActiveSheet();
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
var oDocContent = oShape.GetContent();
var oParagraph = oDocContent.GetElement(0);
oParagraph.AddText("This is a paragraph with the indent of 1 inch set to the first line. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.");
oParagraph.SetIndFirstLine(1440);
var nIndFirstLine = oParagraph.GetIndFirstLine();
oParagraph = Api.CreateParagraph();
oParagraph.AddText("First line indent: " + nIndFirstLine);
oDocContent.Push(oParagraph);
```