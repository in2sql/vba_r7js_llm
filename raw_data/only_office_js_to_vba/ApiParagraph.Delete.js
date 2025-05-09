### Description / Описание
This script adds a flow chart shape to the active worksheet, modifies its content by adding and then deleting a paragraph, and finally updates cell A9 with a confirmation message.
Этот скрипт добавляет форму блок-схемы на активный лист, изменяет её содержимое, добавляя и затем удаляя абзац, и в конце обновляет ячейку A9 сообщением подтверждения.

```vba
' VBA Code
Sub ModifyShapeContent()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Create a solid fill with RGB color
    Dim oFill As FillFormat
    Set oFill = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, _
        60, 35, 200, 100).Fill
    oFill.ForeColor.RGB = RGB(255, 111, 61)
    
    ' Remove the stroke
    With oWorksheet.Shapes(1).Line
        .Weight = 0
        .Visible = msoFalse
    End With
    
    ' Get the shape and its text frame
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes(1)
    
    ' Clear all text
    oShape.TextFrame.Characters.Text = ""
    
    ' Add a new paragraph with text
    oShape.TextFrame.Characters.Text = "This is just a sample text."
    
    ' Delete the paragraph
    oShape.TextFrame.Characters.Text = ""
    
    ' Set value in cell A9
    oWorksheet.Range("A9").Value = "The paragraph from the shape content was removed."
End Sub
```

```javascript
// JavaScript Code
// This script adds a flow chart shape, modifies its content by adding and deleting a paragraph, and updates cell A9.
function modifyShapeContent() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Create a solid fill with RGB color
    var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
    
    // Create no stroke
    var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    
    // Add a shape to the worksheet
    var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
    
    // Get the content of the shape
    var oDocContent = oShape.GetContent();
    
    // Remove all existing elements
    oDocContent.RemoveAllElements();
    
    // Create a new paragraph and add text
    var oParagraph = Api.CreateParagraph();
    oParagraph.AddText("This is just a sample text.");
    
    // Push the paragraph to the shape's content
    oDocContent.Push(oParagraph);
    
    // Delete the paragraph
    oParagraph.Delete();
    
    // Set value in cell A9
    oWorksheet.GetRange("A9").SetValue("The paragraph from the shape content was removed.");
}
```