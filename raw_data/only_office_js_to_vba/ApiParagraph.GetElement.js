```plaintext
// Description
// English: This code adds a shape to the active worksheet, sets its fill and stroke, adds a paragraph with multiple text runs, and applies bold formatting to the third run.
// Russian: Этот код добавляет фигуру на активный лист, устанавливает ее заливку и обводку, добавляет абзац с несколькими текстовыми фрагментами и применяет жирное форматирование к третьему фрагменту.

' VBA Code
Sub AddShapeAndFormatText()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Add shape to worksheet
    ' OnlyOffice shape type "flowChartOnlineStorage" corresponds to msoShapeFlowchartOfflineStorage
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartOfflineStorage, 120, 70, 120, 70)
    
    ' Set the fill color to RGB(255, 111, 61)
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61) ' RGB color
        .Solid
    End With
    
    ' Remove the stroke (no fill)
    With oShape.Line
        .Visible = msoFalse
    End With
    
    ' Clear any existing text
    oShape.TextFrame2.TextRange.Text = ""
    
    ' Add text runs
    With oShape.TextFrame2.TextRange
        ' First text run
        .Text = "This is the text for the first text run. Do not forget a space at its end to separate from the second one. "
        
        ' Second text run
        .InsertAfter "This is the text for the second run. We will set it bold afterwards. It also needs space at its end. "
        
        ' Third text run
        .InsertAfter "This is the text for the third run. It ends the paragraph."
        
        ' Set the third run to bold
        ' Calculate the start position of the third run
        Dim startPos As Long
        startPos = Len("This is the text for the first text run. Do not forget a space at its end to separate from the second one. ") + _
                   Len("This is the text for the second run. We will set it bold afterwards. It also needs space at its end. ") + 1
        ' Set bold for the third run
        .Characters(startPos, Len("This is the text for the third run. It ends the paragraph.")).Font.Bold = msoTrue
    End With
End Sub
```

```javascript
// Description
// English: This code adds a shape to the active worksheet, sets its fill and stroke, adds a paragraph with multiple text runs, and applies bold formatting to the third run.
// Russian: Этот код добавляет фигуру на активный лист, устанавливает ее заливку и обводку, добавляет абзац с несколькими текстовыми фрагментами и применяет жирное форматирование к третьему фрагменту.

// JavaScript Code
function addShapeAndFormatText() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Create solid fill with RGB color
    var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
    
    // Create no stroke
    var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    
    // Add shape to worksheet
    var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
    
    // Get the document content of the shape
    var oDocContent = oShape.GetContent();
    
    // Get the first paragraph element
    var oParagraph = oDocContent.GetElement(0);
    
    // Remove all existing elements in the paragraph
    oParagraph.RemoveAllElements();
    
    // Create first text run and add to paragraph
    var oRun = Api.CreateRun();
    oRun.AddText("This is the text for the first text run. Do not forget a space at its end to separate from the second one. ");
    oParagraph.AddElement(oRun);
    
    // Create second text run and add to paragraph
    oRun = Api.CreateRun();
    oRun.AddText("This is the text for the second run. We will set it bold afterwards. It also needs space at its end. ");
    oParagraph.AddElement(oRun);
    
    // Create third text run and add to paragraph
    oRun = Api.CreateRun();
    oRun.AddText("This is the text for the third run. It ends the paragraph.");
    oParagraph.AddElement(oRun);
    
    // Get the third run and set it to bold
    oRun = oParagraph.GetElement(2);
    oRun.SetBold(true);
}
```