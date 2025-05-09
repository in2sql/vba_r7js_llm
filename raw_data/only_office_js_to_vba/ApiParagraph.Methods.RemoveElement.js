## Description / Описание

**English:**  
This script adds a flowchart shape to the active worksheet, inserts multiple paragraphs with text, and then removes a specific paragraph element.

**Russian:**  
Этот скрипт добавляет фигуру блок-схемы на активный лист, вставляет несколько абзацев с текстом, а затем удаляет конкретный элемент абзаца.

```vba
' VBA Code to add a flowchart shape, insert paragraphs, and remove a specific paragraph element

Sub ModifyShape()
    Dim oWorksheet As Worksheet
    Dim oShape As Shape
    Dim oTextFrame As TextFrame2
    Dim oTextRange As TextRange2
    
    ' Set the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Add a flowchart shape to the worksheet
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, _
        120 * 72 / 25.4, 70 * 72 / 25.4, 200, 100) ' Convert from OnlyOffice units to points
    
    ' Access the text frame of the shape
    Set oTextFrame = oShape.TextFrame2
    
    ' Clear any existing text
    oTextFrame.TextRange.Text = ""
    
    ' Add the first paragraph
    oTextFrame.TextRange.Text = "This is the first paragraph element. "
    
    ' Append the second paragraph
    oTextFrame.TextRange.Text = oTextFrame.TextRange.Text & "This is the second paragraph element. "
    
    ' Append the third paragraph
    oTextFrame.TextRange.Text = oTextFrame.TextRange.Text & "This is the third paragraph element (it will be removed from the paragraph and we will not see it). "
    
    ' Add a line break and fourth paragraph
    oTextFrame.TextRange.Text = oTextFrame.TextRange.Text & vbCrLf & "This is the fourth paragraph element - it became the third, because we removed the previous run from the paragraph. "
    
    ' Add another line break and fifth paragraph
    oTextFrame.TextRange.Text = oTextFrame.TextRange.Text & vbCrLf & "Please note that line breaks are not counted into paragraph elements!"
    
    ' Remove the third paragraph
    Set oTextRange = oTextFrame.TextRange.Paragraphs(3)
    oTextRange.Delete
End Sub
```

```javascript
// JavaScript Code to add a flowchart shape, insert paragraphs, and remove a specific paragraph element

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill color
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a flowchart shape to the worksheet with specified dimensions and styles
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape's document
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Remove all existing elements in the paragraph
oParagraph.RemoveAllElements();

// Create and add the first run of text
var oRun = Api.CreateRun();
oRun.AddText("This is the first paragraph element. ");
oParagraph.AddElement(oRun);

// Create and add the second run of text
oRun = Api.CreateRun();
oRun.AddText("This is the second paragraph element. ");
oParagraph.AddElement(oRun);

// Create and add the third run of text
oRun = Api.CreateRun();
oRun.AddText("This is the third paragraph element (it will be removed from the paragraph and we will not see it). ");
oParagraph.AddElement(oRun);

// Add a line break
oParagraph.AddLineBreak();

// Create and add the fourth run of text
oRun = Api.CreateRun();
oRun.AddText("This is the fourth paragraph element - it became the third, because we removed the previous run from the paragraph. ");
oParagraph.AddElement(oRun);

// Add another line break
oParagraph.AddLineBreak();

// Create and add the fifth run of text
oRun = Api.CreateRun();
oRun.AddText("Please note that line breaks are not counted into paragraph elements!");
oParagraph.AddElement(oRun);

// Remove the third paragraph element
oParagraph.RemoveElement(3);
```