## Description / Описание

**English:**  
This code creates a shape in the active worksheet, adds text to its paragraphs with centered justification, and displays the justification type.

**Russian:**  
Этот код создает фигуру в активном листе, добавляет текст к ее абзацам с выравниванием по центру и отображает тип выравнивания.

```javascript
// This example shows how to get the paragraph contents justification.

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a flow chart online storage shape to the worksheet with specified dimensions and styles
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape's document
var oDocContent = oShape.GetContent();

// Retrieve the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Add text to the paragraph
oParagraph.AddText("This is a paragraph with the text in it aligned by the center. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.");

// Set paragraph justification to center
oParagraph.SetJc("center");

// Get the current justification of the paragraph
var sJc = oParagraph.GetJc();

// Create a new paragraph
oParagraph = Api.CreateParagraph();

// Add text indicating the justification type
oParagraph.AddText("Justification: " + sJc);

// Push the new paragraph to the document content
oDocContent.Push(oParagraph); 
```

```vba
' This macro creates a shape in the active worksheet, adds centered text to it,
' and displays the justification type in a textbox below the shape.

Sub AddShapeWithCenteredText()
    ' Get the active worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Add a flowchart online storage shape to the worksheet with specified position and size
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(Type:=msoShapeFlowchartOnlineStorage, _
                                 Left:=120, Top:=70, Width:=200, Height:=150)
    
    ' Set the fill color to RGB(255, 111, 61)
    shp.Fill.ForeColor.RGB = RGB(255, 111, 61)
    
    ' Remove the shape's outline
    shp.Line.Visible = msoFalse
    
    ' Add text to the shape
    With shp.TextFrame2.TextRange
        .Text = "This is a paragraph with the text in it aligned by the center." & vbCrLf & _
                "These sentences are used to add lines for demonstrative purposes." & vbCrLf & _
                "These sentences are used to add lines for demonstrative purposes."
        ' Set paragraph alignment to center
        .ParagraphFormat.Alignment = msoAlignCenter
    End With
    
    ' Define the justification type
    Dim alignment As String
    alignment = "center"
    
    ' Add a textbox below the shape to display the justification type
    Dim txtBox As Shape
    Set txtBox = ws.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
                                      Left:=120, Top:=230, Width:=300, Height:=50)
    With txtBox.TextFrame2.TextRange
        .Text = "Justification: " & alignment
        ' Set paragraph alignment to left
        .ParagraphFormat.Alignment = msoAlignLeft
    End With
End Sub
```