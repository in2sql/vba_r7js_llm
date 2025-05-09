# Description
This code adds a flowchart shape to the active worksheet, populates it with multiple paragraphs of text, removes a specified paragraph, and adds a concluding paragraph.
Этот код добавляет фигуру блок-схемы на активный лист, заполняет её несколькими абзацами текста, удаляет определённый абзац и добавляет завершающий абзац.

```vba
' VBA Code

Sub ModifyShape()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Create a solid fill color (RGB: 255, 111, 61)
    Dim fillColor As Long
    fillColor = RGB(255, 111, 61)

    ' Add a flowchart shape with specified dimensions
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartOnlineStorage, 200, 60, 200, 100)
    shp.Fill.ForeColor.RGB = fillColor
    shp.Line.Visible = msoFalse

    ' Add the first paragraph of text
    shp.TextFrame.Characters.Text = "This is paragraph #1."

    ' Add additional paragraphs
    Dim i As Integer
    For i = 2 To 5
        shp.TextFrame.Characters.Text = shp.TextFrame.Characters.Text & vbCrLf & "This is paragraph #" & i & "."
    Next i

    ' Split the text into paragraphs
    Dim paras() As String
    paras = Split(shp.TextFrame.Characters.Text, vbCrLf)

    ' Remove the third paragraph if it exists
    If UBound(paras) >= 2 Then
        paras(2) = ""
    End If

    ' Reconstruct the text without the third paragraph
    Dim newText As String
    newText = ""
    For i = LBound(paras) To UBound(paras)
        If paras(i) <> "" Then
            If newText <> "" Then newText = newText & vbCrLf
            newText = newText & paras(i)
        End If
    Next i

    shp.TextFrame.Characters.Text = newText

    ' Add the concluding paragraph
    shp.TextFrame.Characters.Text = shp.TextFrame.Characters.Text & vbCrLf & "We removed paragraph #3, check that out above."
End Sub
```

```javascript
// OnlyOffice JS Code

// This example removes an element using the position specified.
// Этот пример удаляет элемент, используя указанную позицию.

var oWorksheet = Api.GetActiveSheet();

// Create a solid fill color with RGB values (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a flowchart shape with specified parameters
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 200 * 36000, 60 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape's document
var oDocContent = oShape.GetContent();

// Get the first paragraph
var oParagraph = oDocContent.GetElement(0);

// Add text to the first paragraph
oParagraph.AddText("This is paragraph #1.");

// Add additional paragraphs
for (let nParaIncrease = 1; nParaIncrease < 5; ++nParaIncrease) {
    oParagraph = Api.CreateParagraph();
    oParagraph.AddText("This is paragraph #" + (nParaIncrease + 1) + ".");
    oDocContent.Push(oParagraph);
}

// Remove the third paragraph
oDocContent.RemoveElement(2);

// Add a concluding paragraph
oParagraph = Api.CreateParagraph();
oParagraph.AddText("We removed paragraph #3, check that out above.");
oDocContent.Push(oParagraph); 
```