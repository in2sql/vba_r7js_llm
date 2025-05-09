**Setting Paragraph Line Spacing in OnlyOffice using JavaScript and Excel VBA  
Установка межстрочного интервала абзаца в OnlyOffice с использованием JavaScript и Excel VBA**

```javascript
// This example shows how to get and set the paragraph line spacing value.
// Создание фигуры на активном листе и настройка межстрочного интервала абзаца.

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill with RGB color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
// Add a shape to the worksheet with specified dimensions and styling
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
var oDocContent = oShape.GetContent(); // Get the content of the shape
var oParagraph = oDocContent.GetElement(0); // Get the first paragraph element
var oParaPr = oParagraph.GetParaPr(); // Get paragraph properties
oParaPr.SetSpacingLine(3 * 240, "auto"); // Set line spacing to 3 times the normal spacing
// Add text to the paragraph
oParagraph.AddText("Paragraph 1. Spacing: 3 times of a common paragraph line spacing.");
oParagraph.AddLineBreak(); // Add a line break
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.");
var nSpacingLineValue = oParaPr.GetSpacingLineValue(); // Get the current line spacing value
oParagraph = Api.CreateParagraph(); // Create a new paragraph
oParagraph.AddText("Spacing line value : " + nSpacingLineValue); // Add text showing the line spacing value
oDocContent.Push(oParagraph); // Add the new paragraph to the document content
```

```vba
' This example shows how to get and set the paragraph line spacing value.
' Создание фигуры на активном листе и настройка межстрочного интервала абзаца.

Sub SetParagraphLineSpacing()
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet ' Get the active worksheet
    
    ' Add a shape to the worksheet
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, _
                                           120, 70, 200, 150) ' Position and size of the shape
    
    ' Set the fill color of the shape
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61) ' RGB color
        .Solid
    End With
    
    ' Set the line properties of the shape
    With oShape.Line
        .Visible = msoTrue
        .Weight = 0 ' No weight
        .ForeColor.RGB = RGB(255, 255, 255) ' No fill (white)
    End With
    
    ' Add text to the shape
    With oShape.TextFrame2.TextRange
        .Text = "Paragraph 1. Spacing: 3 times of a common paragraph line spacing." & vbCrLf & _
                "These sentences are used to add lines for demonstrative purposes."
        ' Set line spacing to 3 times the normal spacing
        .ParagraphFormat.SpaceWithin = 3
    End With
    
    ' Retrieve the current line spacing value
    Dim nSpacingLineValue As Single
    nSpacingLineValue = oShape.TextFrame2.TextRange.ParagraphFormat.SpaceWithin
    
    ' Add a new paragraph with the spacing line value
    Dim newShape As Shape
    Set newShape = oWorksheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 120, 230, 400, 50)
    With newShape.TextFrame2.TextRange
        .Text = "Spacing line value : " & nSpacingLineValue
    End With
End Sub
```