**Description / Описание**

*English:* This example removes all the elements from the current document or from the current document content.

*Russian:* Этот пример удаляет все элементы из текущего документа или из содержимого текущего документа.

```javascript
// OnlyOffice JS code
// This example removes all the elements from the current document or from the current document content.
var oWorksheet = Api.GetActiveSheet();
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
var oDocContent = oShape.GetContent();
var oParagraph = oDocContent.GetElement(0);
oParagraph.AddText("This is just a sample paragraph.");
oDocContent.RemoveAllElements();
oParagraph = Api.CreateParagraph();
oParagraph.SetJc("left");
oParagraph.AddText("We removed all elements from the shape and added a new paragraph inside it.");
oDocContent.Push(oParagraph); 
```

```vba
' Excel VBA code
' This example removes all the elements from the current worksheet or from the current worksheet content.

Sub RemoveAndAddParagraph()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Create a solid fill color (RGB: 255, 111, 61)
    Dim oFill As FillFormat
    Set oFill = oWorksheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 100, 100).Fill
    oFill.ForeColor.RGB = RGB(255, 111, 61)
    
    ' Create a stroke with no fill
    Dim oStroke As LineFormat
    Set oStroke = oWorksheet.Shapes(1).Line
    oStroke.Visible = msoFalse
    
    ' Add a flowchart shape to the worksheet
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartProcess, 60 * 36000, 35 * 36000, 200, 150)
    oShape.Fill.ForeColor.RGB = RGB(255, 111, 61)
    oShape.Line.Visible = msoFalse
    
    ' Get the text frame of the shape
    Dim oTextFrame As TextFrame2
    Set oTextFrame = oShape.TextFrame2
    
    ' Add text to the first paragraph
    oTextFrame.TextRange.Text = "This is just a sample paragraph."
    
    ' Remove all elements (clear text)
    oTextFrame.TextRange.Text = ""
    
    ' Add a new paragraph with left alignment
    With oTextFrame.TextRange
        .ParagraphFormat.Alignment = msoAlignLeft
        .Text = "We removed all elements from the shape and added a new paragraph inside it."
    End With
End Sub
```