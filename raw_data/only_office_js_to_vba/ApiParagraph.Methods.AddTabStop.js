### Description / Описание

**English:**  
This code adds a shape to the active worksheet, sets its fill color and stroke, and inserts a paragraph with sample text followed by three tab stops. After the tab stops, additional text is added.

**Russian:**  
Этот код добавляет форму на активный лист, устанавливает цвет заливки и обводки, а также вставляет абзац с примерным текстом, за которым следуют три табуляции. После табуляций добавляется дополнительный текст.

```vba
' VBA Code Equivalent

Sub AddShapeWithTabStops()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Define the fill color (RGB: 255, 111, 61)
    Dim fillColor As Long
    fillColor = RGB(255, 111, 61)
    
    ' Add a rectangle shape to the worksheet
    ' Note: Excel VBA uses points for size and position
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeRectangle, 120, 70, 200, 100)
    
    ' Set the fill color
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = fillColor
        .Solid
    End With
    
    ' Set the line (stroke) to no fill
    With oShape.Line
        .Visible = msoTrue
        .Weight = 0
        .ForeColor.RGB = RGB(255, 255, 255) ' White color as no fill
    End With
    
    ' Add text to the shape
    With oShape.TextFrame2
        .TextRange.Text = "This is just a sample text. After it three tab stops will be added." & vbTab & vbTab & vbTab & "This is the text which starts after the tab stops."
        
        ' Set tab stops (in points)
        Dim para As Object
        Set para = .TextRange.Paragraphs(1)
        para.ParagraphFormat.TabStops.ClearAll()
        para.ParagraphFormat.TabStops.Add Position:=100, Alignment:=msoTabStopLeft, Leader:=msoTabLeaderSpaces
        para.ParagraphFormat.TabStops.Add Position:=200, Alignment:=msoTabStopLeft, Leader:=msoTabLeaderSpaces
        para.ParagraphFormat.TabStops.Add Position:=300, Alignment:=msoTabStopLeft, Leader:=msoTabLeaderSpaces
    End With
End Sub
```

```javascript
// OnlyOffice JS Code Equivalent

// This code adds a shape to the active worksheet, sets its fill color and stroke,
// and inserts a paragraph with sample text followed by three tab stops.
// After the tab stops, additional text is added.

function addShapeWithTabStops(Api) {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Create a solid fill with RGB color (255, 111, 61)
    var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
    
    // Create a stroke with no fill
    var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    
    // Add a shape to the worksheet with specified parameters
    var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
    
    // Get the content of the shape
    var oDocContent = oShape.GetContent();
    
    // Get the first paragraph element
    var oParagraph = oDocContent.GetElement(0);
    
    // Add initial text to the paragraph
    oParagraph.AddText("This is just a sample text. After it three tab stops will be added.");
    
    // Add three tab stops
    oParagraph.AddTabStop();
    oParagraph.AddTabStop();
    oParagraph.AddTabStop();
    
    // Add additional text after the tab stops
    oParagraph.AddText("This is the text which starts after the tab stops.");
}
```