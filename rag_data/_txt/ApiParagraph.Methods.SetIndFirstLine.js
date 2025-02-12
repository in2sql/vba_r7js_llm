**Description / Описание**

This script adds a shape to the active worksheet, sets its fill and stroke properties, and manipulates paragraph text with specific indentation settings.

Этот скрипт добавляет фигуру на активный лист, устанавливает ее свойства заливки и обводки, а также обрабатывает текст параграфов с определенными настройками отступа.

---

```vba
' VBA Code to Add and Configure a Shape with Paragraphs in Excel

Sub AddConfiguredShape()
    Dim oWorksheet As Worksheet
    Dim oShape As Shape
    Dim oTextFrame As TextFrame
    Dim oTextRange As TextRange
    
    ' Get the active sheet
    Set oWorksheet = ActiveSheet
    
    ' Add a flowchart shape to the worksheet
    Set oShape = oWorksheet.Shapes.AddShape(Type:=msoShapeFlowchartOnlineStorage, _
                                           Left:=120 * 36, Top:=70 * 36, _
                                           Width:=2 * 36, Height:=3 * 36)
    
    ' Set the fill color to RGB(255, 111, 61)
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61)
        .Solid
    End With
    
    ' Remove the stroke
    With oShape.Line
        .Visible = msoFalse
    End With
    
    ' Access the text frame of the shape
    Set oTextFrame = oShape.TextFrame
    oTextFrame.HorizontalAlignment = xlHAlignLeft
    
    ' Add text with first line indent of 1 inch (72 points)
    Set oTextRange = oTextFrame.Characters
    oTextRange.Text = "This is a paragraph with the indent of 1 inch set to the first line. " & _
                      "These sentences are used to add lines for demonstrative purposes. " & _
                      "These sentences are used to add lines for demonstrative purposes. " & _
                      "These sentences are used to add lines for demonstrative purposes."
    oTextFrame.MarginLeft = 0
    oTextFrame.MarginRight = 0
    oTextFrame.MarginTop = 0
    oTextFrame.MarginBottom = 0
    oTextFrame.Orientation = msoTextOrientationHorizontal
    oTextRange.ParagraphFormat.FirstLineIndent = 72 ' 1 inch
    
    ' Add another paragraph without indentation
    oTextRange.InsertAfter vbCrLf & "This is a paragraph without any indent set to the first line. " & _
                             "These sentences are used to add lines for demonstrative purposes. " & _
                             "These sentences are used to add lines for demonstrative purposes."
    oTextRange.ParagraphFormat.FirstLineIndent = 0
End Sub
```

```javascript
// JavaScript Code to Add and Configure a Shape with Paragraphs in OnlyOffice

// This example sets the paragraph first line indentation.
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a flowchart shape to the worksheet with specified dimensions and styles
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 
                                  120 * 36000, // Left position
                                  70 * 36000,  // Top position
                                  oFill,       // Fill style
                                  oStroke,     // Stroke style
                                  0,           // Rotation
                                  2 * 36000,   // Width
                                  0,           // Height
                                  3 * 36000);  // Depth

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Add text to the paragraph
oParagraph.AddText("This is a paragraph with the indent of 1 inch set to the first line. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.");

// Set the first line indentation to 1440 (assuming units are in twips)
oParagraph.SetIndFirstLine(1440);

// Create a new paragraph
oParagraph = Api.CreateParagraph();

// Add text to the new paragraph without indentation
oParagraph.AddText("This is a paragraph without any indent set to the first line. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.");

// Push the new paragraph to the document content
oDocContent.Push(oParagraph);
```