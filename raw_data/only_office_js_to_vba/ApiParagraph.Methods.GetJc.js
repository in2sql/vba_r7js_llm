**Description / Описание**

This code demonstrates how to create a shape in the active worksheet, add paragraphs with centered justification, and retrieve the justification setting using OnlyOffice API in JavaScript and its equivalent in Excel VBA.

Этот код демонстрирует, как создать фигуру на активном листе, добавить абзацы с выравниванием по центру и получить настройку выравнивания, используя OnlyOffice API в JavaScript и его эквивалент в Excel VBA.

---

```javascript
// JavaScript code using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified dimensions and styles
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Add text to the paragraph
oParagraph.AddText("This is a paragraph with the text in it aligned by the center. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes.");

// Set the justification of the paragraph to center
oParagraph.SetJc("center");

// Retrieve the justification setting
var sJc = oParagraph.GetJc();

// Create a new paragraph
oParagraph = Api.CreateParagraph();

// Add text displaying the justification setting
oParagraph.AddText("Justification: " + sJc);

// Push the new paragraph to the document content
oDocContent.Push(oParagraph);
```

```vba
' Excel VBA equivalent code

Sub CreateShapeAndSetJustification()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet

    ' Define the RGB color
    Dim red As Long, green As Long, blue As Long
    red = 255
    green = 111
    blue = 61
    
    ' Add a shape to the worksheet
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowChartProcess, 120, 70, 200, 150)
    
    ' Set the fill color
    With oShape.Fill
        .ForeColor.RGB = RGB(red, green, blue)
        .Solid
    End With
    
    ' Set the line (stroke) to no fill
    With oShape.Line
        .Visible = msoFalse
    End With
    
    ' Add text to the shape
    With oShape.TextFrame.Characters
        .Text = "This is a paragraph with the text in it aligned by the center. " & _
                "These sentences are used to add lines for demonstrative purposes. " & _
                "These sentences are used to add lines for demonstrative purposes."
        .ParagraphFormat.Alignment = xlCenter
    End With
    
    ' Retrieve the justification setting
    Dim sJc As String
    sJc = "center" ' Since we set it to center
    
    ' Add a new shape to display the justification
    Dim oNewShape As Shape
    Set oNewShape = oWorksheet.Shapes.AddShape(msoShapeRectangle, 350, 70, 300, 50)
    
    ' Set no fill and no line for the new shape
    With oNewShape
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
    End With
    
    ' Add text to the new shape
    With oNewShape.TextFrame.Characters
        .Text = "Justification: " & sJc
        .ParagraphFormat.Alignment = xlLeft
    End With
End Sub
```