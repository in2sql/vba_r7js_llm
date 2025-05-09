### This code demonstrates how to add, group, ungroup shapes, and modify their text in a worksheet.  
### Этот код демонстрирует, как добавлять, группировать, разгруппировать фигуры и изменять их текст на листе.

```javascript
// Get the active worksheet
let oWorksheet = Api.GetActiveSheet();

// Create two solid fill colors
let oFill1 = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
let oFill2 = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));

// Create a stroke with no fill
let oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add the first shape to the worksheet
let oShape1 = oWorksheet.AddShape(
    "flowChartOnlineStorage", 
    60 * 36000, 
    35 * 36000, 
    oFill1, 
    oStroke, 
    0, 
    2 * 36000, 
    0, 
    3 * 36000
);

// Add the second shape to the worksheet
let oShape2 = oWorksheet.AddShape(
    "flowChartOnlineStorage", 
    60 * 36000, 
    35 * 36000, 
    oFill2, 
    oStroke, 
    0, 
    15 * 36000, 
    0, 
    30 * 36000
);

// Group the two shapes
let oGroup = oWorksheet.GroupDrawings([oShape1, oShape2]);

// Ungroup the grouped shapes
oGroup.Ungroup();

// Get the document content of the first shape
let oDocContent1 = oShape1.GetDocContent();

// Get the first paragraph of the document content
let oParagraph1 = oDocContent1.GetElement(0);

// Add text to the first paragraph
oParagraph1.AddText("Shapes are ungrouped");

// Get the document content of the second shape
let oDocContent2 = oShape2.GetDocContent();

// Get the first paragraph of the document content
let oParagraph2 = oDocContent2.GetElement(0);

// Add text to the first paragraph
oParagraph2.AddText("Shapes are ungrouped");
```

```vba
' This VBA code demonstrates how to add, group, ungroup shapes, and modify their text in a worksheet.
' Этот VBA код демонстрирует, как добавлять, группировать, разгруппировать фигуры и изменять их текст на листе.

Sub ManipulateShapes()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' Get the active worksheet
    
    ' Define colors using RGB
    Dim fill1 As Long
    Dim fill2 As Long
    fill1 = RGB(255, 111, 61) ' First fill color
    fill2 = RGB(51, 51, 51)    ' Second fill color
    
    ' Add the first shape to the worksheet
    Dim shape1 As Shape
    Set shape1 = ws.Shapes.AddShape(msoShapeFlowchartOnlineStorage, _
                                    60, 35, 200, 150) ' Adjust size and position as needed
    shape1.Fill.ForeColor.RGB = fill1 ' Set fill color
    shape1.Line.Visible = msoFalse   ' Set no stroke
    
    ' Add the second shape to the worksheet
    Dim shape2 As Shape
    Set shape2 = ws.Shapes.AddShape(msoShapeFlowchartOnlineStorage, _
                                    60, 35, 200, 150) ' Adjust size and position as needed
    shape2.Fill.ForeColor.RGB = fill2 ' Set fill color
    shape2.Line.Visible = msoFalse   ' Set no stroke
    
    ' Group the two shapes
    Dim shapeGroup As Shape
    Set shapeGroup = ws.Shapes.Range(Array(shape1.Name, shape2.Name)).Group
    
    ' Ungroup the grouped shapes
    shapeGroup.Ungroup
    
    ' Add text to the first shape
    shape1.TextFrame.Characters.Text = "Shapes are ungrouped"
    
    ' Add text to the second shape
    shape2.TextFrame.Characters.Text = "Shapes are ungrouped"
End Sub
```