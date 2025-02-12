### Description / Описание
This code demonstrates how to add a shape to the active worksheet, insert two paragraphs into the shape's content, and then make the second paragraph bold.
Этот код демонстрирует, как добавить фигуру на активный лист, вставить два абзаца в содержимое фигуры, а затем сделать второй абзац жирным.

```vba
' VBA Code Equivalent

Sub AddShapeWithParagraphs()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet

    ' Create a fill with RGB color (255, 111, 61)
    Dim oFill As Object
    Set oFill = CreateSolidFill(RGB(255, 111, 61))

    ' Create a stroke with no fill
    Dim oStroke As Object
    Set oStroke = CreateStroke(0, False)

    ' Add a shape to the worksheet
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartStorage, _
        60, 35, 2, 3) ' Dimensions are simplified for VBA

    ' Apply fill and stroke to the shape
    With oShape
        .Fill.ForeColor.RGB = RGB(255, 111, 61)
        .Line.Visible = msoFalse
    End With

    ' Clear any existing text
    oShape.TextFrame2.TextRange.Text = ""

    ' Add the first paragraph
    Dim oParagraph1 As TextRange2
    Set oParagraph1 = oShape.TextFrame2.TextRange.Paragraphs.Add
    oParagraph1.Text = "This is the first paragraph."

    ' Add the second paragraph
    Dim oParagraph2 As TextRange2
    Set oParagraph2 = oShape.TextFrame2.TextRange.Paragraphs.Add
    oParagraph2.Text = "This is the second paragraph."

    ' Make the second paragraph bold
    oParagraph2.Font.Bold = msoTrue
End Sub

' Helper function to create a solid fill (placeholder, as VBA handles fills differently)
Function CreateSolidFill(color As Long) As Object
    ' Implementation depends on specific requirements
    Set CreateSolidFill = Nothing
End Function

' Helper function to create a stroke (placeholder)
Function CreateStroke(weight As Single, visible As Boolean) As Object
    ' Implementation depends on specific requirements
    Set CreateStroke = Nothing
End Function
```

```javascript
// OnlyOffice JS Code

// This example shows how to get the next paragraph.
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
var oNextParagraph = oParagraph1.GetNext();
oNextParagraph.SetBold(true); 
```