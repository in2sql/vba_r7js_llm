```vb
' This example gets a class type and inserts it into the document.
' Этот пример получает тип класса и вставляет его в документ.

Sub InsertShapeAndParagraph()
    ' Get the active worksheet
    Dim oWorksheet As Object
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Create a solid fill with RGB color (255, 111, 61)
    Dim oFill As Object
    Set oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61))
    
    ' Create a stroke with width 0 and no fill
    Dim oStroke As Object
    Set oStroke = Api.CreateStroke(0, Api.CreateNoFill())
    
    ' Add a shape to the worksheet
    Dim oShape As Object
    Set oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000)
    
    ' Get the content of the shape
    Dim oDocContent As Object
    Set oDocContent = oShape.GetContent()
    
    ' Get the first paragraph
    Dim oParagraph As Object
    Set oParagraph = oDocContent.GetElement(0)
    
    ' Get paragraph properties
    Dim oParaPr As Object
    Set oParaPr = oParagraph.GetParaPr()
    
    ' Get the class type of the paragraph
    Dim sClassType As String
    sClassType = oParaPr.GetClassType()
    
    ' Set the first line indent to 1440 (1 inch)
    oParaPr.SetIndFirstLine(1440)
    
    ' Add text to the paragraph
    oParagraph.AddText "This is the first paragraph with the indent of 1 inch set to the first line. "
    oParagraph.AddText "This indent is set by the paragraph style. No paragraph inline style is applied. "
    oParagraph.AddText "These sentences are used to add lines for demonstrative purposes. "
    oParagraph.AddText "These sentences are used to add lines for demonstrative purposes."
    
    ' Create a new paragraph
    Dim oNewParagraph As Object
    Set oNewParagraph = Api.CreateParagraph()
    
    ' Add text with class type to the new paragraph
    oNewParagraph.AddText "Class Type = " & sClassType
    
    ' Push the new paragraph into the document content
    oDocContent.Push oNewParagraph
End Sub
```

```javascript
// This example gets a class type and inserts it into the document.
// Этот пример получает тип класса и вставляет его в документ.

function insertShapeAndParagraph() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Create a solid fill with RGB color (255, 111, 61)
    var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
    
    // Create a stroke with width 0 and no fill
    var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    
    // Add a shape to the worksheet
    var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
    
    // Get the content of the shape
    var oDocContent = oShape.GetContent();
    
    // Get the first paragraph
    var oParagraph = oDocContent.GetElement(0);
    
    // Get paragraph properties
    var oParaPr = oParagraph.GetParaPr();
    
    // Get the class type of the paragraph
    var sClassType = oParaPr.GetClassType();
    
    // Set the first line indent to 1440 (1 inch)
    oParaPr.SetIndFirstLine(1440);
    
    // Add text to the paragraph
    oParagraph.AddText("This is the first paragraph with the indent of 1 inch set to the first line. ");
    oParagraph.AddText("This indent is set by the paragraph style. No paragraph inline style is applied. ");
    oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
    oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.");
    
    // Create a new paragraph
    oParagraph = Api.CreateParagraph();
    
    // Add text with class type to the new paragraph
    oParagraph.AddText("Class Type = " + sClassType);
    
    // Push the new paragraph into the document content
    oDocContent.Push(oParagraph); 
}
```