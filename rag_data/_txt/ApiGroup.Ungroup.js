# Code Description / Описание кода

**English:** This code demonstrates how to ungroup grouped drawing elements in an OnlyOffice worksheet. It creates two shapes, groups them, then ungroups and adds text indicating that the shapes are ungrouped.

**Russian:** Этот код демонстрирует, как разгруппировать сгруппированные элементы рисования на листе OnlyOffice. Он создаёт две фигуры, группирует их, затем разгруппирует и добавляет текст, указывающий, что фигуры разгруппированы.

```vba
' VBA Code
' This code demonstrates how to ungroup grouped drawing elements in an Excel worksheet.

Sub UngroupShapes()
    Dim oWorksheet As Worksheet
    Dim oShape1 As Shape
    Dim oShape2 As Shape
    Dim oGroup As ShapeRange
    
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Add first shape
    Set oShape1 = oWorksheet.Shapes.AddShape(msoShapeFlowchartOfflineStorage, 60, 35, 100, 100)
    oShape1.Fill.ForeColor.RGB = RGB(255, 111, 61) ' Set fill color
    oShape1.Line.Visible = msoFalse ' Set no stroke
    
    ' Add second shape
    Set oShape2 = oWorksheet.Shapes.AddShape(msoShapeFlowchartOfflineStorage, 60, 35, 150, 200)
    oShape2.Fill.ForeColor.RGB = RGB(51, 51, 51) ' Set fill color
    oShape2.Line.Visible = msoFalse ' Set no stroke
    
    ' Group the shapes
    Set oGroup = oWorksheet.Shapes.Range(Array(oShape1.Name, oShape2.Name)).Group
    
    ' Ungroup the shapes
    oGroup.Ungroup
    
    ' Add text to first shape
    oShape1.TextFrame.Characters.Text = "Shapes are ungrouped"
    
    ' Add text to second shape
    oShape2.TextFrame.Characters.Text = "Shapes are ungrouped"
End Sub
```

```javascript
// JavaScript Code
// This code demonstrates how to ungroup grouped drawing elements in an OnlyOffice worksheet.

// Get the active worksheet
let oWorksheet = Api.GetActiveSheet();

// Create solid fills with specified RGB colors
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

// Ungroup the shapes
oGroup.Ungroup();

// Add text to the first shape
let oDocContent1 = oShape1.GetDocContent();
let oParagraph1 = oDocContent1.GetElement(0);
oParagraph1.AddText("Shapes are ungrouped");

// Add text to the second shape
let oDocContent2 = oShape2.GetDocContent();
let oParagraph2 = oDocContent2.GetElement(0);
oParagraph2.AddText("Shapes are ungrouped");
```