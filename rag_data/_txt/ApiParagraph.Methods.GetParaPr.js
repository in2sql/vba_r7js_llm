**Description / Описание:**

This code demonstrates how to add a shape to the active worksheet, set paragraph properties such as spacing after, and add text to the shape using JavaScript for OnlyOffice and equivalent VBA code for Excel.

Этот код демонстрирует, как добавить фигуру на активный лист, установить свойства абзаца, такие как отступ после, и добавить текст в фигуру, используя JavaScript для OnlyOffice и эквивалентный VBA-код для Excel.

```vba
' VBA Code to add a shape, set paragraph spacing, and add text to it
Sub AddShapeAndSetParagraphProperties()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Create a solid fill with RGB color (255, 111, 61)
    Dim fillColor As Long
    fillColor = RGB(255, 111, 61)
    
    ' Add a shape to the worksheet
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartOnlineStorage, 120, 70, 200, 100)
    
    ' Set the fill and line properties
    With shp
        .Fill.ForeColor.RGB = fillColor
        .Line.Visible = msoFalse
    End With
    
    ' Add text to the shape
    shp.TextFrame.Characters.Text = "This is an example of setting a space after a paragraph. " & _
        "The second paragraph will have an offset of one inch from the top. " & _
        "This is due to the fact that the first paragraph has this offset enabled." & vbCrLf & _
        "This is the second paragraph and it is one inch away from the first paragraph."
    
    ' Set paragraph spacing after for the first paragraph
    With shp.TextFrame.Characters(Start:=1, Length:=Len("This is an example of setting a space after a paragraph. The second paragraph will have an offset of one inch from the top. This is due to the fact that the first paragraph has this offset enabled.")).ParagraphFormat
        .SpaceAfter = 18 ' Points (1440 twips = 18 points)
    End With
End Sub
```

```javascript
// JavaScript Code for OnlyOffice API to add a shape, set paragraph spacing, and add text
// This example shows how to get the paragraph properties.
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet
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

// Get the first paragraph
var oParagraph = oDocContent.GetElement(0);

// Get paragraph properties
var oParaPr = oParagraph.GetParaPr();

// Set spacing after to 1440 twips
oParaPr.SetSpacingAfter(1440);

// Add text to the first paragraph
oParagraph.AddText("This is an example of setting a space after a paragraph. ");
oParagraph.AddText("The second paragraph will have an offset of one inch from the top. ");
oParagraph.AddText("This is due to the fact that the first paragraph has this offset enabled.");

// Create a second paragraph
oParagraph = Api.CreateParagraph();
oParagraph.AddText("This is the second paragraph and it is one inch away from the first paragraph.");

// Add the second paragraph to the document content
oDocContent.Push(oParagraph);
```