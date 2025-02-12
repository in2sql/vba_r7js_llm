**Description / Описание**

English: This code creates a shape in the active worksheet, adds two paragraphs of text with specific alignment and indentation using OnlyOffice API and its equivalent implementation in Excel VBA.

Russian: Этот код создает фигуру на активном листе, добавляет два абзаца текста с определенным выравниванием и отступом, используя OnlyOffice API и его эквивалент в Excel VBA.

```javascript
// English: This example sets the paragraph right side indentation.
// Russian: Этот пример устанавливает отступ справа в абзаце.

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill with RGB color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create stroke with no fill
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000); // Add a shape
var oDocContent = oShape.GetContent(); // Get content of the shape
var oParagraph = oDocContent.GetElement(0); // Get the first paragraph
oParagraph.AddText("This is a paragraph with the right offset of 2 inches set to it. "); // Add text
oParagraph.AddText("We also aligned the text in it by the right side. "); // Add text
oParagraph.AddText("This sentence is used to add lines for demonstrative purposes."); // Add text
oParagraph.SetJc("right"); // Set justification to right
oParagraph.SetIndRight(2880); // Set right indentation
oParagraph = Api.CreateParagraph(); // Create a new paragraph
oParagraph.AddText("This is a paragraph without any offset set to it. "); // Add text
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. "); // Add text
oDocContent.Push(oParagraph); // Push the new paragraph to document content
```

```vba
' English: This example sets the paragraph right side indentation.
' Russian: Этот пример устанавливает отступ справа в абзаце.

Sub SetParagraphIndentation()
    Dim oWorksheet As Object
    Set oWorksheet = Api.GetActiveSheet() ' Get the active worksheet
    
    Dim oFill As Object
    Set oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)) ' Create a solid fill with RGB color
    
    Dim oStroke As Object
    Set oStroke = Api.CreateStroke(0, Api.CreateNoFill()) ' Create stroke with no fill
    
    Dim oShape As Object
    Set oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000) ' Add a shape
    
    Dim oDocContent As Object
    Set oDocContent = oShape.GetContent() ' Get content of the shape
    
    Dim oParagraph As Object
    Set oParagraph = oDocContent.GetElement(0) ' Get the first paragraph
    
    oParagraph.AddText "This is a paragraph with the right offset of 2 inches set to it. " ' Add text
    oParagraph.AddText "We also aligned the text in it by the right side. " ' Add text
    oParagraph.AddText "This sentence is used to add lines for demonstrative purposes." ' Add text
    
    oParagraph.SetJc "right" ' Set justification to right
    oParagraph.SetIndRight 2880 ' Set right indentation
    
    Set oParagraph = Api.CreateParagraph() ' Create a new paragraph
    
    oParagraph.AddText "This is a paragraph without any offset set to it. " ' Add text
    oParagraph.AddText "These sentences are used to add lines for demonstrative purposes. " ' Add text
    
    oDocContent.Push oParagraph ' Push the new paragraph to document content
End Sub
```