**Description / Описание:**
This code creates a new shape with specific fill and stroke properties, adds it to the active worksheet, and inserts formatted text into the shape's content.

Этот код создает новую фигуру с определенными свойствами заливки и обводки, добавляет ее на активный лист и вставляет отформатированный текст в содержимое фигуры.

```vba
' VBA Code to create a shape with formatted text in Excel

Sub CreateFormattedShape()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Create RGB color for fill (255, 111, 61)
    Dim fillColor As Long
    fillColor = RGB(255, 111, 61)
    
    ' Add a shape to the worksheet
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, 120, 70, 200, 100)
    
    ' Set the fill color
    oShape.Fill.ForeColor.RGB = fillColor
    
    ' Remove the stroke
    oShape.Line.Visible = msoFalse
    
    ' Add text to the shape
    With oShape.TextFrame.Characters
        .Text = "This is just a sample text. "
        .Font.Name = "Calibri"
    End With
    
    ' Add another text run with a different font
    With oShape.TextFrame.Characters(Start:=Len(oShape.TextFrame.Characters.Text) + 1, Length:= _
        Len("This is a text run with the font family set to 'Comic Sans MS'."))        
        .Text = "This is a text run with the font family set to 'Comic Sans MS'."
        .Font.Name = "Comic Sans MS"
    End With
End Sub
```

```javascript
// JavaScript Code to create a shape with formatted text in OnlyOffice

// This example creates a new smaller text block to be inserted to the paragraph or table.
// Create and configure the shape on the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add the shape to the worksheet with specified dimensions and styles
var oShape = oWorksheet.AddShape(
    "flowChartOnlineStorage", 
    120 * 36000, 
    70 * 36000, 
    oFill, 
    oStroke, 
    0, 
    2 * 36000, 
    0, 
    3 * 36000
);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph in the shape's content
var oParagraph = oDocContent.GetElement(0);

// Create a new text run and add sample text
var oRun = Api.CreateRun();
oRun.AddText("This is just a sample text. ");
oParagraph.AddElement(oRun);

// Create another text run with a different font family
oRun = Api.CreateRun();
oRun.SetFontFamily("Comic Sans MS");
oRun.AddText("This is a text run with the font family set to 'Comic Sans MS'.");
oParagraph.AddElement(oRun);
```