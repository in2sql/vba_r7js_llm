**Description:**
This script adds a flowchart shape to the active worksheet, sets the first line indentation of the first paragraph, and adds multiple lines of text to the shape. It then retrieves and displays the first line indentation value.

**Описание:**
Этот скрипт добавляет фигуру блок-схемы на активный лист, устанавливает отступ первой строки первого абзаца и добавляет несколько строк текста в фигуру. Затем он получает и отображает значение отступа первой строки.

```vba
' VBA Code Equivalent

Sub AddFlowChartShape()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Create a solid fill with RGB color (255, 111, 61)
    Dim oFill As Object
    Set oFill = CreateObject("stdole.StdPicture") ' Placeholder for fill object
    
    ' Create a stroke with no fill
    Dim oStroke As Object
    Set oStroke = Nothing ' No fill
    
    ' Add a flowchart shape to the worksheet
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartPredefined, _
        120, 70, 120, 70) ' Positions and size in points
    
    ' Set the fill color
    oShape.Fill.ForeColor.RGB = RGB(255, 111, 61)
    
    ' Set the stroke (no fill)
    oShape.Line.Visible = msoFalse
    
    ' Access the text frame of the shape
    With oShape.TextFrame2
        ' Get the first paragraph
        Dim oParagraph As TextRange2
        Set oParagraph = .TextRange.Paragraphs(1)
        
        ' Set first line indent (e.g., 1440 which is 1 inch)
        oParagraph.ParagraphFormat.FirstLineIndent = 1440 ' Measurement unit depends on context
        
        ' Add text to the paragraph
        oParagraph.Text = "This is the first paragraph with the indent of 1 inch set to the first line. " & _
                          "This indent is set by the paragraph style. No paragraph inline style is applied. " & _
                          "These sentences are used to add lines for demonstrative purposes. " & _
                          "These sentences are used to add lines for demonstrative purposes. " & _
                          "These sentences are used to add lines for demonstrative purposes."
                          
        ' Retrieve the first line indent value
        Dim nIndFirstLine As Single
        nIndFirstLine = oParagraph.ParagraphFormat.FirstLineIndent
        
        ' Add a new paragraph with the indent value
        Dim oNewParagraph As TextRange2
        Set oNewParagraph = .TextRange.InsertAfter(vbCrLf & "First line indent: " & nIndFirstLine)
    End With
End Sub
```

```javascript
// JavaScript Code Equivalent

// This example shows how to get the paragraph first line indentation.

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a flowchart shape to the worksheet with specified dimensions and styles
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Get the paragraph properties
var oParaPr = oParagraph.GetParaPr();

// Set the first line indentation to 1440 units
oParaPr.SetIndFirstLine(1440);

// Add multiple lines of text to the paragraph
oParagraph.AddText("This is the first paragraph with the indent of 1 inch set to the first line. ");
oParagraph.AddText("This indent is set by the paragraph style. No paragraph inline style is applied. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.");

// Retrieve the first line indentation value
var nIndFirstLine = oParaPr.GetIndFirstLine();

// Create a new paragraph
oParagraph = Api.CreateParagraph();

// Add text displaying the first line indentation
oParagraph.AddText("First line indent: " + nIndFirstLine);

// Push the new paragraph to the document content
oDocContent.Push(oParagraph);
```