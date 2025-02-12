# Description
This code demonstrates how to create and manipulate shapes and paragraphs in OnlyOffice using JavaScript and their VBA equivalent. It adds a shape, inserts paragraphs with specific spacing, and retrieves spacing information.

Этот код демонстрирует создание и манипулирование фигурами и абзацами в OnlyOffice с использованием JavaScript и их эквивалента на VBA. Он добавляет фигуру, вставляет абзацы с определённым отступом и получает информацию об отступе.

```javascript
// Get the active sheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with width 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Add text to the paragraph
oParagraph.AddText("This is an example of setting a space before a paragraph.");
oParagraph.AddText("The second paragraph will have an offset of one inch from the top. ");
oParagraph.AddText("This is due to the fact that the second paragraph has this offset enabled.");

// Create a second paragraph
var oParagraph2 = Api.CreateParagraph();

// Add text to the second paragraph
oParagraph2.AddText("This is the second paragraph and it is one inch away from the first paragraph.");

// Set spacing before the second paragraph
oParagraph2.SetSpacingBefore(1440);

// Push the second paragraph to the document content
oDocContent.Push(oParagraph2);

// Get the spacing before value of the second paragraph
var nSpacingBefore = oParagraph2.GetSpacingBefore();

// Create a new paragraph showing the spacing before
oParagraph = Api.CreateParagraph();
oParagraph.AddText("Spacing before: " + nSpacingBefore);

// Push the new paragraph to the document content
oDocContent.Push(oParagraph);
```

```vba
' Description:
' This VBA code demonstrates how to create and manipulate shapes and paragraphs in Excel. 
' It adds a shape, inserts paragraphs with specific spacing, and retrieves spacing information.

Sub AddShapeAndParagraphs()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Define RGB color
    Dim fillColor As Long
    fillColor = RGB(255, 111, 61)
    
    ' Add a shape to the worksheet
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartStorage, 120, 70, 200, 100)
    
    ' Set the fill color of the shape
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = fillColor
        .Solid
    End With
    
    ' Set the line (stroke) properties
    With oShape.Line
        .Visible = msoFalse ' No fill
    End With
    
    ' Add text to the shape
    With oShape.TextFrame
        .Characters.Text = "This is an example of setting a space before a paragraph." & vbCrLf & _
                           "The second paragraph will have an offset of one inch from the top. " & vbCrLf & _
                           "This is due to the fact that the second paragraph has this offset enabled."
    End With
    
    ' Add a second paragraph with simulated spacing before
    ' Note: Excel VBA does not support paragraph spacing directly, so using line breaks to simulate
    With oShape.TextFrame
        .Characters.Text = .Characters.Text & vbCrLf & _
                           vbCrLf & _
                           "This is the second paragraph and it is one inch away from the first paragraph."
    End With
    
    ' Simulate retrieving spacing before value
    Dim nSpacingBefore As Long
    nSpacingBefore = 1440 ' This value is simulated as Excel does not support direct spacing retrieval
    
    ' Add a new line showing the spacing before
    With oShape.TextFrame
        .Characters.Text = .Characters.Text & vbCrLf & "Spacing before: " & nSpacingBefore
    End With
End Sub
```