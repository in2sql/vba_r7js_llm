**Description / Описание**

This script demonstrates how to manipulate paragraph spacing and add shapes with specific styling in OnlyOffice using JavaScript and its equivalent in Excel VBA.

Этот скрипт демонстрирует, как управлять отступами абзацев и добавлять фигуры с определенным стилем в OnlyOffice с использованием JavaScript и его эквивалента в Excel VBA.

---

```vba
' Excel VBA Equivalent Code
Sub ManipulateParagraphSpacing()
    ' Get the active worksheet
    Dim oWorksheet As Object
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Create a solid fill with RGB color (255, 111, 61)
    Dim oFill As Object
    Set oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61))
    
    ' Create a stroke with no fill
    Dim oStroke As Object
    Set oStroke = Api.CreateStroke(0, Api.CreateNoFill())
    
    ' Add a shape to the worksheet
    Dim oShape As Object
    Set oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000)
    
    ' Get the document content of the shape
    Dim oDocContent As Object
    Set oDocContent = oShape.GetContent()
    
    ' Get the first paragraph element
    Dim oParagraph As Object
    Set oParagraph = oDocContent.GetElement(0)
    
    ' Add text to the first paragraph
    Call oParagraph.AddText("This is an example of setting a space before a paragraph.")
    Call oParagraph.AddText("The second paragraph will have an offset of one inch from the top.")
    Call oParagraph.AddText("This is due to the fact that the second paragraph has this offset enabled.")
    
    ' Create the second paragraph
    Dim oParagraph2 As Object
    Set oParagraph2 = Api.CreateParagraph()
    
    ' Add text to the second paragraph
    Call oParagraph2.AddText("This is the second paragraph and it is one inch away from the first paragraph.")
    
    ' Set spacing before the second paragraph (1440 units)
    Call oParagraph2.SetSpacingBefore(1440)
    
    ' Add the second paragraph to the document content
    Call oDocContent.Push(oParagraph2)
    
    ' Get the spacing before value of the second paragraph
    Dim nSpacingBefore As Long
    nSpacingBefore = oParagraph2.GetSpacingBefore()
    
    ' Create a new paragraph to display the spacing before value
    Set oParagraph = Api.CreateParagraph()
    Call oParagraph.AddText("Spacing before: " & nSpacingBefore)
    
    ' Add the new paragraph to the document content
    Call oDocContent.Push(oParagraph)
End Sub
```

```javascript
// OnlyOffice JavaScript Code Example
// This example shows how to get the spacing before value of the paragraph.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill with RGB color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000); // Add a shape to the worksheet
var oDocContent = oShape.GetContent(); // Get the document content of the shape
var oParagraph = oDocContent.GetElement(0); // Get the first paragraph element
oParagraph.AddText("This is an example of setting a space before a paragraph."); // Add text to the first paragraph
oParagraph.AddText("The second paragraph will have an offset of one inch from the top. "); // Add more text
oParagraph.AddText("This is due to the fact that the second paragraph has this offset enabled."); // Add more text
var oParagraph2 = Api.CreateParagraph(); // Create the second paragraph
oParagraph2.AddText("This is the second paragraph and it is one inch away from the first paragraph."); // Add text to the second paragraph
oParagraph2.SetSpacingBefore(1440); // Set spacing before the second paragraph
oDocContent.Push(oParagraph2); // Add the second paragraph to the document content
var nSpacingBefore = oParagraph2.GetSpacingBefore(); // Get the spacing before value
oParagraph = Api.CreateParagraph(); // Create a new paragraph to display the spacing
oParagraph.AddText("Spacing before: " + nSpacingBefore); // Add text displaying the spacing before value
oDocContent.Push(oParagraph); // Add the new paragraph to the document content
```