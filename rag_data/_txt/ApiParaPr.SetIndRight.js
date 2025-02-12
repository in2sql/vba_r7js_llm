**Description / Описание**

*English:* This script creates a shape on the active worksheet, sets its fill and stroke properties, adds a paragraph to the shape's content, sets the right indentation of the paragraph, and adds multiple lines of text to it.

*Russian:* Этот скрипт создает фигуру на активном листе, устанавливает свойства заливки и обводки, добавляет абзац к содержимому фигуры, устанавливает правый отступ абзаца и добавляет в него несколько строк текста.

```vba
' VBA Code Equivalent

Sub AddShapeWithIndentedParagraph()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Create fill color (RGB: 255, 111, 61)
    Dim oFill As Shape
    Set oFill = oWorksheet.Shapes.AddShape(msoShapeFlowchartDecision, 120, 70, 100, 50)
    oFill.Fill.ForeColor.RGB = RGB(255, 111, 61)
    
    ' Set no stroke
    oFill.Line.Visible = msoFalse
    
    ' Add text to the shape
    With oFill.TextFrame2
        .TextRange.Text = ""
        
        ' Add first paragraph
        .TextRange.Paragraphs(1).ParagraphFormat.RightIndent = 2880 ' Twips
        .TextRange.Text = "This is the first paragraph with the right offset of 2 inches set to it. " & _
                          "This offset is set by the paragraph style. No paragraph inline style is applied. " & _
                          "These sentences are used to add lines for demonstrative purposes."
    End With
End Sub
```

```javascript
// OnlyOffice JS Code Equivalent

// This example sets the paragraph right side indentation.
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified properties
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph in the content
var oParagraph = oDocContent.GetElement(0);

// Get paragraph properties
var oParaPr = oParagraph.GetParaPr();

// Set the right indentation to 2880 (twips)
oParaPr.SetIndRight(2880);

// Add text to the paragraph
oParagraph.AddText("This is the first paragraph with the right offset of 2 inches set to it. ");
oParagraph.AddText("This offset is set by the paragraph style. No paragraph inline style is applied. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
```