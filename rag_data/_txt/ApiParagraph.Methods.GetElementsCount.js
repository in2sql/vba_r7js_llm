# Description / Описание

This code demonstrates how to add a shape to the active worksheet, manipulate its content by removing existing elements, and adding new text with counts of paragraph elements. It creates a solid fill color and a stroke for the shape.  
Этот код демонстрирует, как добавить фигуру на активный лист, изменить ее содержимое путем удаления существующих элементов и добавления нового текста с подсчетом элементов абзаца. Он создает сплошной цвет заливки и обводку для фигуры.

## VBA Code

```vba
' VBA code equivalent to the OnlyOffice JS example
Sub AddShapeAndManipulateContent()
    Dim oWorksheet As Worksheet
    Dim oShape As Shape
    Dim oFillColor As Long
    Dim oStrokeColor As Long
    Dim oTextFrame As TextFrame
    Dim elementsCount As Integer
    Dim runText As String
    
    ' Get the active sheet
    Set oWorksheet = ActiveSheet
    
    ' Define fill color RGB(255, 111, 61)
    oFillColor = RGB(255, 111, 61)
    
    ' Add a shape to the worksheet (using Flowchart Data shape as example)
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartData, 120, 70, 200, 100)
    
    ' Set the fill color
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = oFillColor
        .Solid
    End With
    
    ' Set the line (stroke) to no fill equivalent by making it transparent
    With oShape.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 255) ' White color
        .Transparency = 1 ' Fully transparent
    End With
    
    ' Access the text frame of the shape
    Set oTextFrame = oShape.TextFrame
    
    ' Remove all existing text
    oTextFrame.Characters.Text = ""
    
    ' Initialize elements count
    elementsCount = 0 ' After removal, assume count is 0
    
    ' Create run text with element count
    runText = "Number of paragraph elements at this point: " & elementsCount & vbTab & vbCrLf
    elementsCount = elementsCount + 1 ' After adding text
    runText = runText & "Number of paragraph elements after we added a text run: " & elementsCount
    
    ' Add text to the text frame
    oTextFrame.Characters.Text = runText
End Sub
```

## JavaScript Code

```javascript
// JavaScript code using OnlyOffice API to add and manipulate a shape

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified position and size
var oShape = oWorksheet.AddShape(
    "flowChartOnlineStorage",
    120 * 36000, // x position
    70 * 36000,  // y position
    oFill,       // fill
    oStroke,     // stroke
    0,           // rotation
    2 * 36000,   // width
    0,           // height ratio
    3 * 36000    // height
);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Remove all elements from the paragraph
oParagraph.RemoveAllElements();

// Create a new text run
var oRun = Api.CreateRun();

// Add text to the run
oRun.AddText("Number of paragraph elements at this point: ");

// Add a tab stop
oRun.AddTabStop();

// Add the elements count
oRun.AddText("" + oParagraph.GetElementsCount());

// Add a line break
oRun.AddLineBreak();

// Add the run to the paragraph
oParagraph.AddElement(oRun);

// Add additional text to the run
oRun.AddText("Number of paragraph elements after we added a text run: ");

// Add a tab stop
oRun.AddTabStop();

// Add the updated elements count
oRun.AddText("" + oParagraph.GetElementsCount());
```