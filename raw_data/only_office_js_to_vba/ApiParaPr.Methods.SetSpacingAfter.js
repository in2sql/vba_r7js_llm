## Description / Описание
This code sets the spacing after the current paragraph and adds two paragraphs with specified formatting.
Этот код устанавливает интервал после текущего абзаца и добавляет два абзаца с указанным форматированием.

```vba
' VBA Code to set spacing after a paragraph and add formatted paragraphs

Sub SetParagraphSpacing()
    ' Get the active sheet
    Dim oWorksheet As Object
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Create a solid fill with RGB color
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
    
    ' Set spacing after the paragraph
    Call oParaPr.SetSpacingAfter(1440)
    
    ' Add text to the first paragraph
    Call oParagraph.AddText("This is an example of setting a space after a paragraph. ")
    Call oParagraph.AddText("The second paragraph will have an offset of one inch from the top. ")
    Call oParagraph.AddText("This is due to the fact that the first paragraph has this offset enabled.")
    
    ' Create a new paragraph
    Set oParagraph = Api.CreateParagraph()
    
    ' Add text to the second paragraph
    Call oParagraph.AddText("This is the second paragraph and it is one inch away from the first paragraph.")
    
    ' Push the new paragraph to the document content
    Call oDocContent.Push(oParagraph)
End Sub
```

```javascript
// JavaScript Code to set spacing after the current paragraph and add formatted paragraphs

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph
var oParagraph = oDocContent.GetElement(0);

// Get paragraph properties
var oParaPr = oParagraph.GetParaPr();

// Set spacing after the paragraph
oParaPr.SetSpacingAfter(1440);

// Add text to the first paragraph
oParagraph.AddText("This is an example of setting a space after a paragraph. ");
oParagraph.AddText("The second paragraph will have an offset of one inch from the top. ");
oParagraph.AddText("This is due to the fact that the first paragraph has this offset enabled.");

// Create a new paragraph
oParagraph = Api.CreateParagraph();

// Add text to the second paragraph
oParagraph.AddText("This is the second paragraph and it is one inch away from the first paragraph.");

// Push the new paragraph to the document content
oDocContent.Push(oParagraph);
```