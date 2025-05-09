**Description (English):**
This code demonstrates how to add a shape to the active worksheet, modify its paragraph content by removing existing elements, and display the number of paragraph elements before and after adding a new text run.

**Описание (на русском):**
Этот код демонстрирует, как добавить фигуру на активный лист, изменить её содержимое путем удаления существующих элементов параграфа и отобразить количество элементов параграфа до и после добавления нового текстового блока.

```vba
' VBA Code
' This code adds a shape to the active worksheet, modifies its paragraph content,
' and displays the number of paragraph elements before and after adding a text run.

Sub AddShapeAndModifyParagraph()
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Create fill color
    Dim oFill As Object
    Set oFill = oWorksheet.Shapes.AddShape(msoShapeRectangle, 120, 70, 360, 300).Fill
    oFill.ForeColor.RGB = RGB(255, 111, 61) ' Set RGB color
    
    ' Create stroke with no fill
    With oWorksheet.Shapes(oWorksheet.Shapes.Count)
        .Line.Weight = 0
        .Line.Visible = msoFalse
    End With
    
    ' Add shape of type flowChartOnlineStorage
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartProcess, 120, 70, 360, 300)
    oShape.Fill.ForeColor.RGB = RGB(255, 111, 61)
    oShape.Line.Visible = msoFalse
    
    ' Modify paragraph content
    Dim oTextFrame As TextFrame2
    Set oTextFrame = oShape.TextFrame2
    
    ' Clear existing text
    oTextFrame.TextRange.Text = ""
    
    ' Add initial text run
    oTextFrame.TextRange.Text = "Number of paragraph elements at this point: " _
        & vbTab & oTextFrame.TextRange.Paragraphs(1).Words.Count & vbNewLine
    
    ' Add new text run
    oTextFrame.TextRange.Text = oTextFrame.TextRange.Text & "Number of paragraph elements after we added a text run: " _
        & vbTab & (oTextFrame.TextRange.Paragraphs(1).Words.Count + 1)
End Sub
```

```javascript
// OnlyOffice JS Code
// This code adds a shape to the active worksheet, modifies its paragraph content,
// and displays the number of paragraph elements before and after adding a text run.

// Get the active sheet
var oWorksheet = Api.GetActiveSheet();

// Create fill color
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add shape to the worksheet
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Remove all elements from the paragraph
oParagraph.RemoveAllElements();

// Create a new run
var oRun = Api.CreateRun();

// Add initial text
oRun.AddText("Number of paragraph elements at this point: ");

// Add tab stop
oRun.AddTabStop();

// Add text with the number of elements
oRun.AddText("" + oParagraph.GetElementsCount());

// Add a line break
oRun.AddLineBreak();

// Add the run to the paragraph
oParagraph.AddElement(oRun);

// Add additional text
oRun.AddText("Number of paragraph elements after we added a text run: ");

// Add another tab stop
oRun.AddTabStop();

// Add text with the updated number of elements
oRun.AddText("" + oParagraph.GetElementsCount());
```