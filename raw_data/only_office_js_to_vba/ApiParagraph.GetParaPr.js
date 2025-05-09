# Code Description / Описание кода

**English**: This code demonstrates how to set paragraph spacing and add text to a shape in OnlyOffice using JavaScript and its equivalent in Excel VBA.

**Russian**: Этот код демонстрирует, как установить отступы для абзаца и добавить текст в фигуру в OnlyOffice, используя JavaScript и его эквивалент в Excel VBA.

```javascript
// JavaScript code
// This example shows how to get the paragraph properties.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000); // Add a shape to the worksheet
var oDocContent = oShape.GetContent(); // Get the content of the shape
var oParagraph = oDocContent.GetElement(0); // Get the first paragraph
var oParaPr = oParagraph.GetParaPr(); // Get paragraph properties
oParaPr.SetSpacingAfter(1440); // Set spacing after paragraph
oParagraph.AddText("This is an example of setting a space after a paragraph. ");
oParagraph.AddText("The second paragraph will have an offset of one inch from the top. ");
oParagraph.AddText("This is due to the fact that the first paragraph has this offset enabled.");
oParagraph = Api.CreateParagraph(); // Create a new paragraph
oParagraph.AddText("This is the second paragraph and it is one inch away from the first paragraph.");
oDocContent.Push(oParagraph); // Add the new paragraph to the content
```

```vba
' VBA code
' This example shows how to set paragraph spacing and add text to a shape in OnlyOffice.

Sub SetParagraphProperties()
    ' Get the active worksheet
    Dim oWorksheet As Object
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Create a solid fill color
    Dim oFill As Object
    Set oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61))
    
    ' Create a stroke with no fill
    Dim oStroke As Object
    Set oStroke = Api.CreateStroke(0, Api.CreateNoFill())
    
    ' Add a shape to the worksheet
    Dim oShape As Object
    Set oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000)
    
    ' Get the content of the shape
    Dim oDocContent As Object
    Set oDocContent = oShape.GetContent()
    
    ' Get the first paragraph
    Dim oParagraph As Object
    Set oParagraph = oDocContent.GetElement(0)
    
    ' Get paragraph properties
    Dim oParaPr As Object
    Set oParaPr = oParagraph.GetParaPr()
    
    ' Set spacing after paragraph
    oParaPr.SetSpacingAfter 1440
    
    ' Add text to the paragraph
    oParagraph.AddText "This is an example of setting a space after a paragraph. "
    oParagraph.AddText "The second paragraph will have an offset of one inch from the top. "
    oParagraph.AddText "This is due to the fact that the first paragraph has this offset enabled."
    
    ' Create a new paragraph
    Set oParagraph = Api.CreateParagraph()
    
    ' Add text to the new paragraph
    oParagraph.AddText "This is the second paragraph and it is one inch away from the first paragraph."
    
    ' Add the new paragraph to the content
    oDocContent.Push oParagraph
End Sub
```