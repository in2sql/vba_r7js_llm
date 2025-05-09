**Description / Описание:**

This code sets the justification of paragraph content to center and adds text to the paragraph.

Этот код устанавливает выравнивание содержимого параграфа по центру и добавляет текст к параграфу.

```javascript
// OnlyOffice JS Code
// This example sets the paragraph contents justification.

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill with specified RGB color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000); // Add a shape to the worksheet
var oDocContent = oShape.GetContent(); // Get the content of the shape
var oParagraph = oDocContent.GetElement(0); // Get the first paragraph
var oParaPr = oParagraph.GetParaPr(); // Get paragraph properties
oParaPr.SetJc("center"); // Set justification to center
oParagraph.AddText("This is a paragraph with the text in it aligned by the center. ");
oParagraph.AddText("The justification is specified in the paragraph style. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes."); 
```

```vba
' Excel VBA Code
' This code sets the justification of paragraph content to center and adds text to the paragraph.

Sub SetParagraphJustification()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' Get the active worksheet
    
    ' Create a solid fill with RGB color
    Dim fillColor As Long
    fillColor = RGB(255, 111, 61)
    
    ' Create a shape with specified type and position
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartStorageData, 120, 70, 200, 150) ' Adjust position and size as needed
    
    ' Set fill and stroke properties
    With shp
        .Fill.ForeColor.RGB = fillColor ' Set fill color
        .Line.Visible = msoFalse ' No stroke
    End With
    
    ' Add text to the shape and set paragraph alignment
    With shp.TextFrame2.TextRange
        .Text = "This is a paragraph with the text in it aligned by the center. " & _
                "The justification is specified in the paragraph style. " & _
                "These sentences are used to add lines for demonstrative purposes. " & _
                "These sentences are used to add lines for demonstrative purposes."
        .ParagraphFormat.Alignment = msoAlignCenter ' Set justification to center
    End With
End Sub
```