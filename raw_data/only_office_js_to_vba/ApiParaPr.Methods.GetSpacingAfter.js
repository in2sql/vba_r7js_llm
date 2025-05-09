**Description:**

This code demonstrates how to set the spacing after a paragraph and add text to a shape in OnlyOffice using both JavaScript and Excel VBA.

Этот код демонстрирует, как установить интервал после абзаца и добавить текст к фигуре в OnlyOffice с использованием JavaScript и Excel VBA.

```vba
' VBA code equivalent to OnlyOffice JavaScript example
Sub SetParagraphSpacing()
    ' Get the active sheet
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
    
    ' Get paragraph properties
    Dim oParaPr As Object
    Set oParaPr = oParagraph.GetParaPr()
    
    ' Set spacing after the paragraph
    oParaPr.SetSpacingAfter 1440
    
    ' Add text to the paragraph
    oParagraph.AddText "This is an example of setting a space after a paragraph. "
    oParagraph.AddText "The second paragraph will have an offset of one inch from the top. "
    oParagraph.AddText "These sentences are used to add lines for demonstrative purposes. "
    oParagraph.AddText "These sentences are used to add lines for demonstrative purposes. "
    oParagraph.AddText "These sentences are used to add lines for demonstrative purposes."
    
    ' Get the spacing after value
    Dim nSpacingAfter As Long
    nSpacingAfter = oParaPr.GetSpacingAfter()
    
    ' Create a new paragraph and add it to the document content
    Set oParagraph = Api.CreateParagraph()
    oParagraph.AddText "Spacing after : " & nSpacingAfter
    oDocContent.Push oParagraph
End Sub
```

```javascript
// JavaScript code using OnlyOffice API to set paragraph spacing and add text
function setParagraphSpacing() {
    // Get the active sheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Create a solid fill with RGB color (255, 111, 61)
    var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
    
    // Create a stroke with no fill
    var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    
    // Add a shape to the worksheet
    var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
    
    // Get the document content of the shape
    var oDocContent = oShape.GetContent();
    
    // Get the first paragraph element
    var oParagraph = oDocContent.GetElement(0);
    
    // Get paragraph properties
    var oParaPr = oParagraph.GetParaPr();
    
    // Set spacing after the paragraph
    oParaPr.SetSpacingAfter(1440);
    
    // Add text to the paragraph
    oParagraph.AddText("This is an example of setting a space after a paragraph. ");
    oParagraph.AddText("The second paragraph will have an offset of one inch from the top. ");
    oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
    oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
    oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.");
    
    // Get the spacing after value
    var nSpacingAfter = oParaPr.GetSpacingAfter();
    
    // Create a new paragraph and add it to the document content
    oParagraph = Api.CreateParagraph();
    oParagraph.AddText("Spacing after : " + nSpacingAfter);
    oDocContent.Push(oParagraph);
}
```