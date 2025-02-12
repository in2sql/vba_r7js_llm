# Add a Paragraph to a Shape in the Active Worksheet
# Добавление абзаца в фигуру на активном листе

```javascript
// This example adds a paragraph in document content.

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified parameters
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 200 * 36000, 60 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Remove all existing elements from the shape
oDocContent.RemoveAllElements();

// Create a new paragraph
var oParagraph = Api.CreateParagraph();

// Add text to the paragraph
oParagraph.AddText("We removed all elements from the shape and added a new paragraph inside it.");

// Add the paragraph to the document content
oDocContent.AddElement(oParagraph);

// Push the paragraph to the content stack
oDocContent.Push(oParagraph); 
```

```vba
' This example adds a paragraph in document content.
' Этот пример добавляет абзац в содержимое документа.

Sub AddParagraphToShape()
    ' Get the active worksheet
    Dim oWorksheet As Object
    Set oWorksheet = Api.GetActiveSheet
    
    ' Create a solid fill with RGB color (255, 111, 61)
    Dim oFill As Object
    Set oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61))
    
    ' Create a stroke with no fill
    Dim oStroke As Object
    Set oStroke = Api.CreateStroke(0, Api.CreateNoFill())
    
    ' Add a shape to the worksheet with specified parameters
    Dim oShape As Object
    Set oShape = oWorksheet.AddShape("flowChartOnlineStorage", 200 * 36000, 60 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000)
    
    ' Get the content of the shape
    Dim oDocContent As Object
    Set oDocContent = oShape.GetContent
    
    ' Remove all existing elements from the shape
    oDocContent.RemoveAllElements
    
    ' Create a new paragraph
    Dim oParagraph As Object
    Set oParagraph = Api.CreateParagraph
    
    ' Add text to the paragraph
    oParagraph.AddText "We removed all elements from the shape and added a new paragraph inside it."
    
    ' Add the paragraph to the document content
    oDocContent.AddElement oParagraph
    
    ' Push the paragraph to the content stack
    oDocContent.Push oParagraph
End Sub
```