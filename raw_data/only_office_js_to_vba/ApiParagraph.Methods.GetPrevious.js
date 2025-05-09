```plaintext
// This script demonstrates how to add a shape with text to the active worksheet,
// create and manipulate paragraphs within the shape, and format text.
// Этот скрипт демонстрирует, как добавить фигуру с текстом на активный лист,
// создать и изменить абзацы внутри фигуры, а также отформатировать текст.
```

```vba
' VBA Code Equivalent

Sub AddShapeAndFormatText()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Create a shape with specific dimensions and formatting
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartConnector, 60, 35, 200, 100)
    
    ' Set the fill color of the shape
    oShape.Fill.ForeColor.RGB = RGB(255, 111, 61)
    
    ' Remove the border (stroke) of the shape
    oShape.Line.Visible = msoFalse
    
    ' Clear any existing text
    oShape.TextFrame.Characters.Text = ""
    
    ' Add the first paragraph
    oShape.TextFrame.Characters.Text = "This is the first paragraph." & vbCrLf
    
    ' Add the second paragraph
    oShape.TextFrame.Characters.Text = oShape.TextFrame.Characters.Text & "This is the second paragraph."
    
    ' Make the first paragraph bold
    With oShape.TextFrame.Characters(Start:=1, Length:=Len("This is the first paragraph.")).Font
        .Bold = True
    End With
End Sub
```

```javascript
// This example shows how to get the previous paragraph.
var oWorksheet = Api.GetActiveSheet();
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
var oDocContent = oShape.GetContent();
oDocContent.RemoveAllElements();
var oParagraph1 = Api.CreateParagraph();
oParagraph1.AddText("This is the first paragraph.");
oDocContent.Push(oParagraph1);
var oParagraph2 = Api.CreateParagraph();
oParagraph2.AddText("This is the second paragraph.");
oDocContent.Push(oParagraph2);
var oPreviousParagraph = oParagraph2.GetPrevious();
oPreviousParagraph.SetBold(true);
```