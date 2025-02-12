**Description / Описание:**

English: This code sets the first line indentation for a paragraph and adds text to a shape in Excel using VBA, equivalent to the provided OnlyOffice JavaScript code.

Russian: Этот код устанавливает отступ первой строки абзаца и добавляет текст в фигуру в Excel с использованием VBA, аналогично предоставленному коду OnlyOffice на JavaScript.

```vba
' VBA Code to set paragraph first line indentation and add text to a shape

Sub SetParagraphIndent()
    ' Get the active worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Add a flowchart storage shape to the worksheet
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartOnlineStorage, 120, 70, 120, 70)
    
    ' Access the text frame of the shape
    With shp.TextFrame2
        ' Set the first line indent to 1 inch (144 points)
        .TextRange.ParagraphFormat.FirstLineIndent = 144
        ' Add text to the shape
        .TextRange.Text = "This is the first paragraph with the indent of 1 inch set to the first line. " & _
                         "This indent is set by the paragraph style. No paragraph inline style is applied. " & _
                         "These sentences are used to add lines for demonstrative purposes. " & _
                         "These sentences are used to add lines for demonstrative purposes. " & _
                         "These sentences are used to add lines for demonstrative purposes."
    End With
End Sub
```

```javascript
// JavaScript Code to set paragraph first line indentation using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a no-fill stroke
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a flowchart online storage shape to the worksheet with specified dimensions and styles
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

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element in the shape
var oParagraph = oDocContent.GetElement(0);

// Get the paragraph properties
var oParaPr = oParagraph.GetParaPr();

// Set the first line indentation to 1440 (1 inch)
oParaPr.SetIndFirstLine(1440);

// Add text to the paragraph
oParagraph.AddText("This is the first paragraph with the indent of 1 inch set to the first line. ");
oParagraph.AddText("This indent is set by the paragraph style. No paragraph inline style is applied. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.");
```