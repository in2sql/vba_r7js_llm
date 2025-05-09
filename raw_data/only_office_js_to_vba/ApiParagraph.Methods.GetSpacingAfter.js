**Example of setting spacing after paragraphs in OnlyOffice via JavaScript and its equivalent in Excel VBA**  
**Пример установки отступа после абзацев в OnlyOffice с использованием JavaScript и его эквивалент в Excel VBA**

```vba
' VBA code equivalent to OnlyOffice JavaScript example
Sub SetParagraphSpacing()
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
    
    ' Get the content of the shape
    Dim oDocContent As Object
    Set oDocContent = oShape.GetContent()
    
    ' Get the first paragraph
    Dim oParagraph1 As Object
    Set oParagraph1 = oDocContent.GetElement(0)
    
    ' Add text to the first paragraph
    oParagraph1.AddText "This is an example of setting a space after a paragraph. "
    oParagraph1.AddText "The second paragraph will have an offset of one inch from the top. "
    oParagraph1.AddText "This is due to the fact that the first paragraph has this offset enabled."
    
    ' Set spacing after for the first paragraph
    oParagraph1.SetSpacingAfter 1440
    
    ' Create the second paragraph
    Dim oParagraph2 As Object
    Set oParagraph2 = Api.CreateParagraph()
    
    ' Add text to the second paragraph
    oParagraph2.AddText "This is the second paragraph and it is one inch away from the first paragraph."
    
    ' Add a line break
    oParagraph2.AddLineBreak
    
    ' Get spacing after value from the first paragraph
    Dim nSpacingAfter As Long
    nSpacingAfter = oParagraph1.GetSpacingAfter()
    
    ' Add spacing information to the second paragraph
    oParagraph2.AddText "Spacing after: " & nSpacingAfter
    
    ' Push the second paragraph to the document content
    oDocContent.Push oParagraph2
End Sub
```

```javascript
// JavaScript code equivalent to the VBA example

// This example shows how to set the spacing after a paragraph in OnlyOffice
function setParagraphSpacing() {
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
    
    // Get the first paragraph
    var oParagraph1 = oDocContent.GetElement(0);
    
    // Add text to the first paragraph
    oParagraph1.AddText("This is an example of setting a space after a paragraph. ");
    oParagraph1.AddText("The second paragraph will have an offset of one inch from the top. ");
    oParagraph1.AddText("This is due to the fact that the first paragraph has this offset enabled.");
    
    // Set spacing after for the first paragraph
    oParagraph1.SetSpacingAfter(1440);
    
    // Create the second paragraph
    var oParagraph2 = Api.CreateParagraph();
    
    // Add text to the second paragraph
    oParagraph2.AddText("This is the second paragraph and it is one inch away from the first paragraph.");
    
    // Add a line break
    oParagraph2.AddLineBreak();
    
    // Get spacing after value from the first paragraph
    var nSpacingAfter = oParagraph1.GetSpacingAfter();
    
    // Add spacing information to the second paragraph
    oParagraph2.AddText("Spacing after: " + nSpacingAfter);
    
    // Push the second paragraph to the document content
    oDocContent.Push(oParagraph2);
}

// Execute the function to set paragraph spacing
setParagraphSpacing();
```