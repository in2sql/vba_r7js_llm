**Description:**
*English:* This code sets the paragraph line spacing and adds text to a shape in an Excel worksheet.
*Russian:* Этот код устанавливает межстрочный интервал абзаца и добавляет текст в фигуру на листе Excel.

**VBA Code:**
```vba
' This VBA code sets the paragraph line spacing and adds text to a shape in the active worksheet
Sub SetParagraphSpacingAndAddText()
    Dim oWorksheet As Worksheet
    Dim oShape As Shape
    Dim oTextFrame As TextFrame
    Dim oTextRange As TextRange
    Dim oParagraphFormat As ParagraphFormat

    ' Get the active worksheet
    Set oWorksheet = ActiveSheet

    ' Add a shape to the worksheet
    Set oShape = oWorksheet.Shapes.AddShape(Type:=msoShapeFlowchartOnlineStorage, _
                                           Left:=120, Top:=70, Width:=200, Height:=100)
    
    ' Get the text frame of the shape
    Set oTextFrame = oShape.TextFrame
    
    ' Add text to the shape
    oTextFrame.Characters.Text = "Paragraph 1. Spacing: 3 times of a common paragraph line spacing." & vbCrLf & _
                                 "These sentences are used to add lines for demonstrative purposes. " & _
                                 "These sentences are used to add lines for demonstrative purposes."
    
    ' Get the text range
    Set oTextRange = oTextFrame.Characters
    
    ' Get the paragraph format
    Set oParagraphFormat = oTextRange.ParagraphFormat
    
    ' Set the line spacing to 3 times the default
    oParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
    oParagraphFormat.LineSpacing = 36 ' 3 times 12pt
    
End Sub
```

**OnlyOffice JavaScript Code:**
```javascript
// This example sets the paragraph line spacing and adds text to a shape in the active sheet
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

// Get the paragraph properties
var oParaPr = oParagraph.GetParaPr();

// Set the line spacing to 3 times the default with automatic spacing
oParaPr.SetSpacingLine(3 * 240, "auto");

// Add the first line of text
oParagraph.AddText("Paragraph 1. Spacing: 3 times of a common paragraph line spacing.");

// Add a line break
oParagraph.AddLineBreak();

// Add additional text lines
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
```