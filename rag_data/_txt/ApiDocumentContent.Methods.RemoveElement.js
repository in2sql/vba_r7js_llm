```markdown
# Code Description
This script adds a flowchart shape to the active worksheet, populates it with multiple paragraphs of text, removes the third paragraph, and adds a final paragraph indicating the removal.
Этот скрипт добавляет фигуру блок-схемы на активный лист, заполняет ее несколькими абзацами текста, удаляет третий абзац и добавляет последний абзац с указанием удаления.

## Excel VBA Code
```vba
' This VBA script adds a flowchart shape, manipulates its content by adding and removing paragraphs.

Sub ManipulateShapeContent()
    Dim oWorksheet As Object
    Dim oFill As Object
    Dim oStroke As Object
    Dim oShape As Object
    Dim oDocContent As Object
    Dim oParagraph As Object
    Dim nParaIncrease As Integer

    ' Get the active sheet
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Create a solid fill with RGB color (255, 111, 61)
    Set oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61))
    
    ' Create a stroke with width 0 and no fill
    Set oStroke = Api.CreateStroke(0, Api.CreateNoFill())
    
    ' Add a shape of type "flowChartOnlineStorage" with specified dimensions
    Set oShape = oWorksheet.AddShape("flowChartOnlineStorage", 200 * 36000, 60 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000)
    
    ' Get the content of the shape
    Set oDocContent = oShape.GetContent()
    
    ' Get the first paragraph
    Set oParagraph = oDocContent.GetElement(0)
    
    ' Add text to the first paragraph
    Call oParagraph.AddText("This is paragraph #1.")
    
    ' Add paragraphs #2 to #5
    For nParaIncrease = 1 To 4
        Set oParagraph = Api.CreateParagraph()
        Call oParagraph.AddText("This is paragraph #" & (nParaIncrease + 1) & ".")
        Call oDocContent.Push(oParagraph)
    Next nParaIncrease
    
    ' Remove the third paragraph (index 2)
    Call oDocContent.RemoveElement(2)
    
    ' Add a final paragraph indicating removal
    Set oParagraph = Api.CreateParagraph()
    Call oParagraph.AddText("We removed paragraph #3, check that out above.")
    Call oDocContent.Push(oParagraph)
End Sub
```

## OnlyOffice JS Code
```javascript
// This script adds a flowchart shape, manipulates its content by adding and removing paragraphs.

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with width 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape of type "flowChartOnlineStorage" with specified dimensions
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 200 * 36000, 60 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph
var oParagraph = oDocContent.GetElement(0);

// Add text to the first paragraph
oParagraph.AddText("This is paragraph #1.");

// Add paragraphs #2 to #5
for (let nParaIncrease = 1; nParaIncrease < 5; ++nParaIncrease) {
    oParagraph = Api.CreateParagraph();
    oParagraph.AddText("This is paragraph #" + (nParaIncrease + 1) + ".");
    oDocContent.Push(oParagraph);
}

// Remove the third paragraph (index 2)
oDocContent.RemoveElement(2);

// Add a final paragraph indicating removal
oParagraph = Api.CreateParagraph();
oParagraph.AddText("We removed paragraph #3, check that out above.");
oDocContent.Push(oParagraph);
```