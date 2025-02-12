**Description / Описание**

This code sets the paragraph line spacing in a shape on the active worksheet.
Этот код устанавливает межстрочный интервал абзаца в фигуре на активном листе.

```javascript
// This example sets the paragraph line spacing.
var oWorksheet = Api.GetActiveSheet();
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
var oDocContent = oShape.GetContent();
var oParagraph = oDocContent.GetElement(0);
var oParaPr = oParagraph.GetParaPr();
oParaPr.SetSpacingLine(3 * 240, "auto");
oParagraph.AddText("Paragraph 1. Spacing: 3 times of a common paragraph line spacing.");
oParagraph.AddLineBreak();
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. "); 
```

```vba
' This example sets the paragraph line spacing in a shape on the active worksheet
Sub SetParagraphLineSpacing()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim txtRange As TextRange2
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Add a shape with specified parameters
    ' Parameters: type, left, top, width, height
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartProcess, 120, 70, 200, 150)
    
    ' Set the fill color to RGB(255, 111, 61)
    shp.Fill.ForeColor.RGB = RGB(255, 111, 61)
    
    ' Remove the stroke
    shp.Line.Visible = msoFalse
    
    ' Access the text frame of the shape
    Set txtRange = shp.TextFrame2.TextRange
    
    ' Set the paragraph line spacing to 3 times the default
    With txtRange.ParagraphFormat
        .LineSpacingRule = msoLineSpacingMultiple
        .LineSpacing = 3 * 12 ' Assuming 12 points as the base line spacing
    End With
    
    ' Add text to the shape
    txtRange.Text = "Paragraph 1. Spacing: 3 times of a common paragraph line spacing." & vbCrLf & _
                   "These sentences are used to add lines for demonstrative purposes. " & _
                   "These sentences are used to add lines for demonstrative purposes."
End Sub
```