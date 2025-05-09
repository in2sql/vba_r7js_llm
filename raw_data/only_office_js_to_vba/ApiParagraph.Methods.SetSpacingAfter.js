**Description / Описание**

This code sets the spacing after a paragraph and adds a second paragraph with a one-inch offset.

Этот код устанавливает отступ после абзаца и добавляет второй абзац с отступом в один дюйм.

```vba
' VBA Code to set spacing after a paragraph and add a second paragraph with an offset

Sub SetParagraphSpacing()
    Dim oWorksheet As Object
    Dim oFill As Object
    Dim oStroke As Object
    Dim oShape As Object
    Dim oDocContent As Object
    Dim oParagraph As Object
    
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
    
    ' Get the first paragraph element
    Set oParagraph = oDocContent.GetElement(0)
    
    ' Add text to the first paragraph
    Call oParagraph.AddText("This is an example of setting a space after a paragraph. ")
    Call oParagraph.AddText("The second paragraph will have an offset of one inch from the top. ")
    Call oParagraph.AddText("This is due to the fact that the first paragraph has this offset enabled.")
    
    ' Set spacing after the first paragraph (1440 twips = 1 inch)
    Call oParagraph.SetSpacingAfter(1440)
    
    ' Create a second paragraph
    Set oParagraph = Api.CreateParagraph()
    
    ' Add text to the second paragraph
    Call oParagraph.AddText("This is the second paragraph and it is one inch away from the first paragraph.")
    
    ' Push the second paragraph to the document content
    Call oDocContent.Push(oParagraph)
End Sub
```

```javascript
// JavaScript Code to set spacing after a paragraph and add a second paragraph with an offset

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

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Add text to the first paragraph
oParagraph.AddText("This is an example of setting a space after a paragraph. ");
oParagraph.AddText("The second paragraph will have an offset of one inch from the top. ");
oParagraph.AddText("This is due to the fact that the first paragraph has this offset enabled.");

// Set spacing after the first paragraph (1440 twips = 1 inch)
oParagraph.SetSpacingAfter(1440);

// Create a second paragraph
oParagraph = Api.CreateParagraph();

// Add text to the second paragraph
oParagraph.AddText("This is the second paragraph and it is one inch away from the first paragraph.");

// Push the second paragraph to the document content
oDocContent.Push(oParagraph);
```