# Description (Описание)

**English:** This script creates a shape in the active worksheet, manipulates its content by adding text runs, removes all elements from a paragraph, and adds a new text run.

**Russian:** Этот скрипт создает фигуру на активном листе, манипулирует ее содержимым, добавляя текстовые элементы, удаляет все элементы из абзаца и добавляет новый текстовый элемент.

```javascript
// JavaScript Code using OnlyOffice API

// Get the active sheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with width 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified parameters
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Create a new text run
var oRun = Api.CreateRun();

// Add text to the run
oRun.AddText("This is the first text run in the current paragraph.");

// Add the run to the paragraph
oParagraph.AddElement(oRun);

// Remove all elements from the paragraph
oParagraph.RemoveAllElements();

// Create another text run
oRun = Api.CreateRun();

// Add new text to the run
oRun.AddText("We removed all the paragraph elements and added a new text run inside it.");

// Add the new run to the paragraph
oParagraph.AddElement(oRun);

// Push the modified paragraph back to the document content
oDocContent.Push(oParagraph);
```

```vba
' VBA Code equivalent to the OnlyOffice API JavaScript code

Sub ManipulateShapeContent()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Add a shape to the worksheet
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowChartProcess, 120, 70, 200, 100)
    
    ' Set the fill color to RGB(255, 111, 61)
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61)
        .Solid
    End With
    
    ' Remove the line (stroke)
    With oShape.Line
        .Visible = msoFalse
    End With
    
    ' Access the text frame of the shape
    With oShape.TextFrame2
        ' Ensure there is a paragraph
        If .TextRange.Paragraphs.Count = 0 Then
            .TextRange.Text = ""
        End If
        
        ' Add the first text run
        .TextRange.Paragraphs(1).Text = "This is the first text run in the current paragraph."
        
        ' Remove all elements from the paragraph
        .TextRange.Paragraphs(1).Text = ""
        
        ' Add a new text run
        .TextRange.Paragraphs(1).Text = "We removed all the paragraph elements and added a new text run inside it."
    End With
End Sub
```