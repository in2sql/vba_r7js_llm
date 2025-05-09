### Description / Описание

**English:**  
This code adds a shape to the active worksheet in OnlyOffice, sets the paragraph line spacing, adds text to the shape, and retrieves the spacing line value.

**Russian:**  
Этот код добавляет фигуру на активный лист в OnlyOffice, устанавливает межстрочный интервал абзаца, добавляет текст в фигуру и извлекает значение межстрочного интервала.

```javascript
// JavaScript code using OnlyOffice API
// This code adds a shape to the active worksheet, sets paragraph line spacing, adds text, and retrieves spacing line value.

var oWorksheet = Api.GetActiveSheet();
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill with RGB color (255, 111, 61)
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with width 0 and no fill
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 80 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000); // Add shape with specified dimensions and styles
var oDocContent = oShape.GetContent(); // Get the content of the shape
var oParagraph = oDocContent.GetElement(0); // Get the first paragraph element
oParagraph.SetSpacingLine(3 * 240, "auto"); // Set line spacing
oParagraph.AddText("Paragraph 1. Spacing: 3 times of a common paragraph line spacing."); // Add text
oParagraph.AddLineBreak(); // Add line break
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. "); // Add text
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. "); // Add text
oParagraph.AddLineBreak(); // Add line break
var nSpacingLineValue = oParagraph.GetSpacingLineValue(); // Get spacing line value
oParagraph.AddText("Spacing line value: " + nSpacingLineValue); // Add text showing spacing line value
```

```vba
' VBA code equivalent for OnlyOffice API
' This code adds a shape to the active worksheet, sets paragraph line spacing, adds text, and retrieves spacing line value.

Sub AddShapeAndSetParagraphSpacing()
    ' Get the active sheet
    Dim oWorksheet As Object
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Create a solid fill with RGB color (255, 111, 61)
    Dim oFill As Object
    Set oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61))
    
    ' Create a stroke with width 0 and no fill
    Dim oStroke As Object
    Set oStroke = Api.CreateStroke(0, Api.CreateNoFill())
    
    ' Add shape with specified dimensions and styles
    Dim oShape As Object
    Set oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 80 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000)
    
    ' Get the content of the shape
    Dim oDocContent As Object
    Set oDocContent = oShape.GetContent()
    
    ' Get the first paragraph element
    Dim oParagraph As Object
    Set oParagraph = oDocContent.GetElement(0)
    
    ' Set line spacing
    oParagraph.SetSpacingLine 3 * 240, "auto"
    
    ' Add text
    oParagraph.AddText "Paragraph 1. Spacing: 3 times of a common paragraph line spacing."
    
    ' Add line break
    oParagraph.AddLineBreak
    
    ' Add more text
    oParagraph.AddText "These sentences are used to add lines for demonstrative purposes. "
    oParagraph.AddText "These sentences are used to add lines for demonstrative purposes. "
    
    ' Add another line break
    oParagraph.AddLineBreak
    
    ' Get spacing line value
    Dim nSpacingLineValue As Variant
    nSpacingLineValue = oParagraph.GetSpacingLineValue()
    
    ' Add text showing spacing line value
    oParagraph.AddText "Spacing line value: " & nSpacingLineValue
End Sub
```