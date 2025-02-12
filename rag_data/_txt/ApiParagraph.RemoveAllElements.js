**Description:**
This VBA and JavaScript code demonstrate how to remove all elements from the current paragraph, add a new text run, and manipulate shape properties in a worksheet.
**Описание:**
Этот VBA и JavaScript код демонстрирует, как удалить все элементы из текущего абзаца, добавить новый текстовый блок и изменить свойства формы на листе.

```javascript
// This example removes all the elements from the current paragraph.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill with specified RGB color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000); // Add a shape to the worksheet
var oDocContent = oShape.GetContent(); // Get the content of the shape
var oParagraph = oDocContent.GetElement(0); // Get the first paragraph element
var oRun = Api.CreateRun(); // Create a new text run
oRun.AddText("This is the first text run in the current paragraph."); // Add text to the run
oParagraph.AddElement(oRun); // Add the run to the paragraph
oParagraph.RemoveAllElements(); // Remove all elements from the paragraph
oRun = Api.CreateRun(); // Create another new text run
oRun.AddText("We removed all the paragraph elements and added a new text run inside it."); // Add new text to the run
oParagraph.AddElement(oRun); // Add the new run to the paragraph
oDocContent.Push(oParagraph); // Update the document content with the modified paragraph
```

```vba
' This example removes all the elements from the current paragraph.
Sub ModifyParagraph()
    Dim oWorksheet As Object
    Set oWorksheet = Api.GetActiveSheet() ' Get the active worksheet
    
    Dim oFill As Object
    Set oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)) ' Create a solid fill with specified RGB color
    
    Dim oStroke As Object
    Set oStroke = Api.CreateStroke(0, Api.CreateNoFill()) ' Create a stroke with no fill
    
    Dim oShape As Object
    Set oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000) ' Add a shape to the worksheet
    
    Dim oDocContent As Object
    Set oDocContent = oShape.GetContent() ' Get the content of the shape
    
    Dim oParagraph As Object
    Set oParagraph = oDocContent.GetElement(0) ' Get the first paragraph element
    
    Dim oRun As Object
    Set oRun = Api.CreateRun() ' Create a new text run
    oRun.AddText "This is the first text run in the current paragraph." ' Add text to the run
    oParagraph.AddElement oRun ' Add the run to the paragraph
    
    oParagraph.RemoveAllElements ' Remove all elements from the paragraph
    
    Set oRun = Api.CreateRun() ' Create another new text run
    oRun.AddText "We removed all the paragraph elements and added a new text run inside it." ' Add new text to the run
    oParagraph.AddElement oRun ' Add the new run to the paragraph
    
    oDocContent.Push oParagraph ' Update the document content with the modified paragraph
End Sub
```