## Description / Описание

**English:**  
This code adds a shape to the active worksheet, sets the left indentation of the first paragraph within the shape, adds text to the paragraph, retrieves the indentation value, and then adds a new paragraph displaying the indentation value.

**Русский:**  
Этот код добавляет фигуру на активный рабочий лист, устанавливает левый отступ первого абзаца внутри фигуры, добавляет текст в абзац, извлекает значение отступа, а затем добавляет новый абзац, отображающий значение отступа.

```vba
' VBA Code Equivalent

Sub AddShapeAndSetIndent()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Add a shape to the worksheet
    ' Parameters: Type, Left, Top, Width, Height
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartPredefined, _
        120, 70, 200, 100) ' Adjust size as needed
    
    ' Set the fill color of the shape
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61)
        .Solid
    End With
    
    ' Access the text frame of the shape
    Dim oTextFrame As TextFrame
    Set oTextFrame = oShape.TextFrame
    
    ' Set the text to have two paragraphs
    oTextFrame.Characters.Text = "This is the first paragraph with the indent of 2 inches set to it. " & _
        "This indent is set by the paragraph style. No paragraph inline style is applied."
    
    ' Set left indent of the first paragraph
    With oTextFrame.Paragraphs(1).ParagraphFormat
        .LeftIndent = Application.InchesToPoints(2) ' 2 inches indent
    End With
    
    ' Get the left indent value
    Dim nIndLeft As Single
    nIndLeft = oTextFrame.Paragraphs(1).ParagraphFormat.LeftIndent
    
    ' Add a new paragraph with the indent value
    oTextFrame.Characters.Text = oTextFrame.Characters.Text & vbCrLf & "Left indent: " & nIndLeft
End Sub
```

```javascript
// OnlyOffice JS Code Equivalent

// This example shows how to get the paragraph left side indentation.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill with specified RGB color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
// Add a shape to the worksheet with specified parameters
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
var oDocContent = oShape.GetContent(); // Get the content of the shape
var oParagraph = oDocContent.GetElement(0); // Get the first paragraph element
var oParaPr = oParagraph.GetParaPr(); // Get the paragraph properties
oParaPr.SetIndLeft(2880); // Set left indentation
// Add text to the paragraph
oParagraph.AddText("This is the first paragraph with the indent of 2 inches set to it. ");
oParagraph.AddText("This indent is set by the paragraph style. No paragraph inline style is applied. ");
var nIndLeft = oParaPr.GetIndLeft(); // Get the left indentation value
oParagraph = Api.CreateParagraph(); // Create a new paragraph
oParagraph.AddText("Left indent: " + nIndLeft); // Add text with the indentation value
oDocContent.Push(oParagraph); // Push the new paragraph to the document content
```