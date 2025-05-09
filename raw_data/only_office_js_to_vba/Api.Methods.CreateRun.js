**Description / Описание**

*English:*  
This code creates a new shape with a specific fill and stroke on the active worksheet. It then adds a paragraph to the shape containing two text runs, one with default font and another with the font family set to 'Comic Sans MS'.

*Русский:*  
Этот код создает новую фигуру с определенной заливкой и обводкой на активном листе. Затем он добавляет абзац к фигуре, содержащий два текстовых элемента, один с шрифтом по умолчанию и другой с установленным семейством шрифтов 'Comic Sans MS'.

---

### VBA Code

```vba
' Create a new shape with specific fill and stroke on the active worksheet
Sub CreateShapeWithText()
    Dim oWorksheet As Object
    Dim oFill As Object
    Dim oStroke As Object
    Dim oShape As Object
    Dim oDocContent As Object
    Dim oParagraph As Object
    Dim oRun As Object

    ' Get the active sheet
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Create a solid fill with RGB color (255, 111, 61)
    Set oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61))
    
    ' Create a stroke with width 0 and no fill
    Set oStroke = Api.CreateStroke(0, Api.CreateNoFill())
    
    ' Add a shape to the worksheet
    Set oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000)
    
    ' Get the document content of the shape
    Set oDocContent = oShape.GetContent()
    
    ' Get the first paragraph
    Set oParagraph = oDocContent.GetElement(0)
    
    ' Create a new text run and add text
    Set oRun = Api.CreateRun()
    oRun.AddText "This is just a sample text. "
    oParagraph.AddElement oRun
    
    ' Create another text run, set font family, and add text
    Set oRun = Api.CreateRun()
    oRun.SetFontFamily "Comic Sans MS"
    oRun.AddText "This is a text run with the font family set to 'Comic Sans MS'."
    oParagraph.AddElement oRun
End Sub
```

### OnlyOffice JS Code

```javascript
// Create a new shape with specific fill and stroke on the active worksheet
function createShapeWithText() {
    // Get the active sheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Create a solid fill with RGB color (255, 111, 61)
    var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
    
    // Create a stroke with width 0 and no fill
    var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    
    // Add a shape to the worksheet
    var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
    
    // Get the document content of the shape
    var oDocContent = oShape.GetContent();
    
    // Get the first paragraph
    var oParagraph = oDocContent.GetElement(0);
    
    // Create a new text run and add text
    var oRun = Api.CreateRun();
    oRun.AddText("This is just a sample text. ");
    oParagraph.AddElement(oRun);
    
    // Create another text run, set font family, and add text
    oRun = Api.CreateRun();
    oRun.SetFontFamily("Comic Sans MS");
    oRun.AddText("This is a text run with the font family set to 'Comic Sans MS'.");
    oParagraph.AddElement(oRun);
}
```