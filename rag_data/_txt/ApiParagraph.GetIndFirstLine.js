### Description
This code creates a shape in the active worksheet, adds a paragraph with specific text, sets the first line indentation, and displays the indentation value.

Этот код создает фигуру на активном листе, добавляет абзац с определенным текстом, устанавливает отступ первой строки и отображает значение отступа.

```vba
' VBA Code Equivalent for OnlyOffice API Example

Sub CreateShapeWithIndentedParagraph()
    Dim oWorksheet As Object
    Dim oFill As Object
    Dim oStroke As Object
    Dim oShape As Object
    Dim oDocContent As Object
    Dim oParagraph As Object
    Dim nIndFirstLine As Long
    
    ' Get the active sheet
    Set oWorksheet = Application.ActiveSheet
    
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
    
    ' Add text to the paragraph
    Call oParagraph.AddText("This is a paragraph with the indent of 1 inch set to the first line. ")
    Call oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ")
    Call oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ")
    Call oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.")
    
    ' Set first line indentation to 1 inch (1440 twips)
    Call oParagraph.SetIndFirstLine(1440)
    
    ' Get the first line indentation value
    nIndFirstLine = oParagraph.GetIndFirstLine()
    
    ' Create a new paragraph
    Set oParagraph = Api.CreateParagraph()
    
    ' Add text displaying the indentation value
    Call oParagraph.AddText("First line indent: " & nIndFirstLine)
    
    ' Push the new paragraph to the document content
    Call oDocContent.Push(oParagraph)
End Sub
```

```javascript
// JavaScript Code Using OnlyOffice API Example

// This example shows how to create a shape, add a paragraph with text,
// set the first line indentation, and display the indentation value.

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

// Add text to the paragraph
oParagraph.AddText("This is a paragraph with the indent of 1 inch set to the first line. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.");

// Set first line indentation to 1 inch (1440 twips)
oParagraph.SetIndFirstLine(1440);

// Get the first line indentation value
var nIndFirstLine = oParagraph.GetIndFirstLine();

// Create a new paragraph
oParagraph = Api.CreateParagraph();

// Add text displaying the indentation value
oParagraph.AddText("First line indent: " + nIndFirstLine);

// Push the new paragraph to the document content
oDocContent.Push(oParagraph);
```