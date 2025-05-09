**English:** This script creates a new paragraph within a shape in the active worksheet by setting fill and stroke properties, removing existing elements, and adding the paragraph.

**Русский:** Этот скрипт создает новый абзац внутри фигуры на активном листе, устанавливая свойства заливки и обводки, удаляя существующие элементы и добавляя абзац.

```javascript
// This example creates a new paragraph.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill with specified RGB color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
// Add a shape with specified parameters to the worksheet
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
var oDocContent = oShape.GetContent(); // Get the content of the shape
oDocContent.RemoveAllElements(); // Remove all existing elements from the shape's content
var oParagraph = Api.CreateParagraph(); // Create a new paragraph
oParagraph.SetJc("left"); // Set paragraph alignment to left
oParagraph.AddText("We removed all elements from the shape and added a new paragraph inside it."); // Add text to the paragraph
oDocContent.Push(oParagraph); // Add the paragraph to the shape's content
```

```vba
' This example creates a new paragraph.
Sub CreateNewParagraph()
    Dim oWorksheet As Object
    Set oWorksheet = Api.GetActiveSheet() ' Get the active worksheet
    
    Dim oFill As Object
    Set oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)) ' Create a solid fill with specified RGB color
    
    Dim oStroke As Object
    Set oStroke = Api.CreateStroke(0, Api.CreateNoFill()) ' Create a stroke with no fill
    
    ' Add a shape with specified parameters to the worksheet
    Dim oShape As Object
    Set oShape = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000)
    
    Dim oDocContent As Object
    Set oDocContent = oShape.GetContent() ' Get the content of the shape
    
    oDocContent.RemoveAllElements ' Remove all existing elements from the shape's content
    
    Dim oParagraph As Object
    Set oParagraph = Api.CreateParagraph() ' Create a new paragraph
    
    oParagraph.SetJc "left" ' Set paragraph alignment to left
    
    oParagraph.AddText "We removed all elements from the shape and added a new paragraph inside it." ' Add text to the paragraph
    
    oDocContent.Push oParagraph ' Add the paragraph to the shape's content
End Sub
```