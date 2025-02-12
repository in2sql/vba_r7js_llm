## Set Paragraph Right Indentation / Установка правого отступа абзаца

**English:** This code sets the paragraph right side indentation and adds text to a shape in Excel VBA and OnlyOffice JavaScript.

**Russian:** Этот код устанавливает правый отступ абзаца и добавляет текст к фигуре в Excel VBA и OnlyOffice JavaScript.

```vba
' Excel VBA Code to set paragraph right indentation and add text to a shape

Sub SetParagraphRightIndentation()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim txtFrame As TextFrame2
    Dim para As TextRange2
    
    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Add a shape to the worksheet
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartProcess, 120, 70, 300, 150)
    
    ' Set the fill color (RGB 255,111,61)
    shp.Fill.ForeColor.RGB = RGB(255, 111, 61)
    
    ' Remove the line (no fill)
    shp.Line.Visible = msoFalse
    
    ' Access the text frame
    Set txtFrame = shp.TextFrame2
    
    ' Access the first paragraph
    Set para = txtFrame.TextRange.Paragraphs(1)
    
    ' Set right indentation to 144 points (2 inches)
    para.ParagraphFormat.RightIndent = 144
    
    ' Add text to the paragraph
    para.Text = "This is the first paragraph with the right offset of 2 inches set to it. " & _
                "This offset is set by the paragraph style. No paragraph inline style is applied. " & _
                "These sentences are used to add lines for demonstrative purposes."
End Sub
```

```javascript
// This example sets the paragraph right side indentation.
// Этот пример устанавливает правый отступ абзаца.

var oWorksheet = Api.GetActiveSheet(); // Get active sheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create fill color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create no border
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000); // Add shape
var oDocContent = oShape.GetContent(); // Get shape content
var oParagraph = oDocContent.GetElement(0); // Get first paragraph
var oParaPr = oParagraph.GetParaPr(); // Get paragraph properties
oParaPr.SetIndRight(2880); // Set right indentation to 2880 twips (2 inches)
oParagraph.AddText("This is the first paragraph with the right offset of 2 inches set to it. ");
oParagraph.AddText("This offset is set by the paragraph style. No paragraph inline style is applied. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
```