**Description / Описание:**

This script demonstrates how to manipulate an Excel worksheet by adding shapes and setting paragraph spacing using OnlyOffice API in JavaScript and its equivalent in Excel VBA.

Этот скрипт демонстрирует, как манипулировать рабочим листом Excel, добавляя фигуры и устанавливая отступы абзацев с использованием OnlyOffice API на JavaScript и его эквивалента в Excel VBA.

---

```vba
' VBA Code to add shapes and set paragraph spacing in Excel

Sub AddShapeAndSetParagraphSpacing()
    Dim oWorksheet As Worksheet
    Dim oShape As Shape
    Dim oParagraph As Object
    Dim oParagraph2 As Object
    Dim oParaPr As Object
    Dim nSpacingBefore As Long
    
    ' Set the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Add a shape to the worksheet
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, _
                                           120 * 72, 70 * 72, 200, 100)
    
    ' Add text to the shape
    oShape.TextFrame.Characters.Text = "This is an example of setting a space before a paragraph." & vbCrLf & _
                                       "The second paragraph will have an offset of one inch from the top." & vbCrLf & _
                                       "This is due to the fact that the second paragraph has this offset enabled."
    
    ' Add a second paragraph
    Set oParagraph2 = oShape.TextFrame2.TextRange.Paragraphs.Add
    oParagraph2.Text = "This is the second paragraph and it is one inch away from the first paragraph."
    
    ' Set spacing before the second paragraph
    oParagraph2.ParagraphFormat.SpaceBefore = 144 ' Points (1 inch = 72 points)
    
    ' Retrieve the spacing before value
    nSpacingBefore = oParagraph2.ParagraphFormat.SpaceBefore
    
    ' Add a paragraph displaying the spacing before
    oShape.TextFrame2.TextRange.Paragraphs.Add.Text = "Spacing before: " & nSpacingBefore
End Sub
```

```javascript
// JavaScript Code using OnlyOffice API to add shapes and set paragraph spacing

function addShapeAndSetParagraphSpacing() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Create a solid fill with RGB color
    var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
    
    // Create a stroke with no fill
    var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    
    // Add a shape to the worksheet
    var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
    
    // Get the content of the shape
    var oDocContent = oShape.GetContent();
    
    // Get the first paragraph
    var oParagraph = oDocContent.GetElement(0);
    
    // Add text to the first paragraph
    oParagraph.AddText("This is an example of setting a space before a paragraph.");
    oParagraph.AddText("The second paragraph will have an offset of one inch from the top. ");
    oParagraph.AddText("This is due to the fact that the second paragraph has this offset enabled.");
    
    // Create the second paragraph
    var oParagraph2 = Api.CreateParagraph();
    oParagraph2.AddText("This is the second paragraph and it is one inch away from the first paragraph.");
    
    // Push the second paragraph to the document content
    oDocContent.Push(oParagraph2);
    
    // Get paragraph properties and set spacing before
    var oParaPr = oParagraph2.GetParaPr();
    oParaPr.SetSpacingBefore(1440); // Assuming 1440 units = 1 inch
    
    // Retrieve the spacing before value
    var nSpacingBefore = oParaPr.GetSpacingBefore();
    
    // Create a new paragraph to display spacing before
    oParagraph = Api.CreateParagraph();
    oParagraph.AddText("Spacing before: " + nSpacingBefore);
    oDocContent.Push(oParagraph);
}
```