# Description / Описание

**English:** This code sets the spacing before the current paragraph and adds a shape to the active worksheet in OnlyOffice. The second paragraph is offset by one inch from the first paragraph.

**Русский:** Этот код устанавливает интервал перед текущим абзацем и добавляет фигуру на активный лист в OnlyOffice. Второй абзац отступает на один дюйм от первого абзаца.

```vba
' VBA code equivalent to the OnlyOffice JS code

Sub SetParagraphSpacingAndAddShape()
    Dim oWorksheet As Object
    Dim oFill As Object
    Dim oStroke As Object
    Dim oShape As Object
    Dim oDocContent As Object
    Dim oParagraph As Object
    Dim oParaPr As Object

    ' Get the active sheet
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Create a solid fill with RGB color (255, 111, 61)
    Set oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61))
    
    ' Create a stroke with no fill
    Set oStroke = Api.CreateStroke(0, Api.CreateNoFill())
    
    ' Add a shape to the worksheet
    Set oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000)
    
    ' Get the content of the shape
    Set oDocContent = oShape.GetContent()
    
    ' Get the first paragraph
    Set oParagraph = oDocContent.GetElement(0)
    
    ' Add text to the paragraph
    oParagraph.AddText "This is an example of setting a space before a paragraph. "
    oParagraph.AddText "The second paragraph will have an offset of one inch from the top. "
    oParagraph.AddText "This is due to the fact that the second paragraph has this offset enabled."
    
    ' Create a new paragraph
    Set oParagraph = Api.CreateParagraph()
    
    ' Get paragraph properties
    Set oParaPr = oParagraph.GetParaPr()
    
    ' Set spacing before to 1440 (twips)
    oParaPr.SetSpacingBefore 1440
    
    ' Add text to the new paragraph
    oParagraph.AddText "This is the second paragraph and it is one inch away from the first paragraph."
    
    ' Push the new paragraph to the document content
    oDocContent.Push oParagraph
End Sub
```

```javascript
// JavaScript code using OnlyOffice API to set paragraph spacing and add a shape

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph
var oParagraph = oDocContent.GetElement(0);

// Add text to the paragraph
oParagraph.AddText("This is an example of setting a space before a paragraph. ");
oParagraph.AddText("The second paragraph will have an offset of one inch from the top. ");
oParagraph.AddText("This is due to the fact that the second paragraph has this offset enabled.");

// Create a new paragraph
oParagraph = Api.CreateParagraph();

// Get paragraph properties
var oParaPr = oParagraph.GetParaPr();

// Set spacing before to 1440 (twips)
oParaPr.SetSpacingBefore(1440);

// Add text to the new paragraph
oParagraph.AddText("This is the second paragraph and it is one inch away from the first paragraph.");

// Push the new paragraph to the document content
oDocContent.Push(oParagraph);
```