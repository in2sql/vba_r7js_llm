```javascript
// Description in English: This code creates a shape with specified fill and stroke, adds paragraphs with line spacing and text, and retrieves the line spacing rule.
// Описание на русском: Этот код создаёт фигуру с заданной заливкой и обводкой, добавляет абзацы с межстрочным интервалом и текстом, а также получает правило межстрочного интервала.

```javascript
// This example shows how to get the paragraph line spacing rule.
var oWorksheet = Api.GetActiveSheet();
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 80 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
var oDocContent = oShape.GetContent();
var oParagraph = oDocContent.GetElement(0);
oParagraph.SetSpacingLine(3 * 240, "auto");
oParagraph.AddText("Paragraph 1. Spacing: 3 times of a common paragraph line spacing.");
oParagraph.AddLineBreak();
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddLineBreak();
var sSpacingLineRule = oParagraph.GetSpacingLineRule();
oParagraph.AddText("Spacing line rule: " + sSpacingLineRule); 
```

```vba
' Description in English: This VBA code creates a shape with specified fill and stroke, adds paragraphs with line spacing and text, and retrieves the line spacing rule.
' Описание на русском: Этот код VBA создаёт фигуру с заданной заливкой и обводкой, добавляет абзацы с межстрочным интервалом и текстом, а также получает правило межстрочного интервала.

Sub CreateShapeAndParagraph()
    ' This code creates a shape with specified fill and stroke, adds paragraphs with line spacing and text, and retrieves the line spacing rule.
    Dim oWorksheet As Object
    Dim oFill As Object
    Dim oStroke As Object
    Dim oShape As Object
    Dim oDocContent As Object
    Dim oParagraph As Object
    Dim sSpacingLineRule As String
    
    ' Get the active sheet
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Create solid fill color
    Set oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61))
    
    ' Create stroke with no fill
    Set oStroke = Api.CreateStroke(0, Api.CreateNoFill())
    
    ' Add shape to the worksheet
    Set oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 80 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000)
    
    ' Get document content of the shape
    Set oDocContent = oShape.GetContent()
    
    ' Get the first paragraph element
    Set oParagraph = oDocContent.GetElement(0)
    
    ' Set line spacing with 3 times and auto rule
    oParagraph.SetSpacingLine 3 * 240, "auto"
    
    ' Add text to the paragraph
    oParagraph.AddText "Paragraph 1. Spacing: 3 times of a common paragraph line spacing."
    oParagraph.AddLineBreak
    oParagraph.AddText "These sentences are used to add lines for demonstrative purposes. "
    oParagraph.AddText "These sentences are used to add lines for demonstrative purposes. "
    oParagraph.AddLineBreak
    
    ' Retrieve the line spacing rule
    sSpacingLineRule = oParagraph.GetSpacingLineRule()
    oParagraph.AddText "Spacing line rule: " & sSpacingLineRule
End Sub
```