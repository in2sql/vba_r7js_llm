**Description:**

- **English:** This code creates bullets for paragraphs with a specified numbering character or symbol.
- **Russian:** Этот код создает маркеры для абзацев с указанным символом или номером.

```vba
' VBA Code to create bullets for paragraphs

Sub CreateBulletedParagraphs()
    Dim oWorksheet As Object
    Dim oFill As Object
    Dim oStroke As Object
    Dim oShape As Object
    Dim oDocContent As Object
    Dim oParagraph As Object
    Dim oBullet As Object

    ' Get the active sheet
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Create a solid fill with RGB color
    Set oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61))
    
    ' Create a stroke with no fill
    Set oStroke = Api.CreateStroke(0, Api.CreateNoFill())
    
    ' Add a shape to the worksheet
    Set oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000)
    
    ' Get the content of the shape
    Set oDocContent = oShape.GetContent()
    
    ' Get the first paragraph element
    Set oParagraph = oDocContent.GetElement(0)
    
    ' Create a numbering bullet
    Set oBullet = Api.CreateNumbering("ArabicParenR", 1)
    
    ' Set the bullet for the paragraph
    oParagraph.SetBullet oBullet
    
    ' Add text to the paragraph
    oParagraph.AddText "This is an example of the numbered paragraph."
    
    ' Create a new paragraph
    Set oParagraph = Api.CreateParagraph()
    
    ' Set the bullet for the new paragraph
    oParagraph.SetBullet oBullet
    
    ' Add text to the new paragraph
    oParagraph.AddText "This is an example of the numbered paragraph."
    
    ' Add the new paragraph to the document content
    oDocContent.Push oParagraph
End Sub
```

```javascript
// JavaScript Code to create bullets for paragraphs

// This example creates a bullet for a paragraph with the numbering character or symbol specified with the sType parameter.
var oWorksheet = Api.GetActiveSheet();
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
var oDocContent = oShape.GetContent();
var oParagraph = oDocContent.GetElement(0);
var oBullet = Api.CreateNumbering("ArabicParenR", 1);
oParagraph.SetBullet(oBullet);
oParagraph.AddText("This is an example of the numbered paragraph.");
oParagraph = Api.CreateParagraph();
oParagraph.SetBullet(oBullet);
oParagraph.AddText("This is an example of the numbered paragraph.");
oDocContent.Push(oParagraph);
```