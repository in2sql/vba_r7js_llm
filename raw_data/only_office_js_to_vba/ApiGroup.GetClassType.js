**Description / Описание**

This code retrieves the active worksheet, creates two shapes with different fill colors, groups them, obtains the class type of the group, and adds text indicating the class type to each shape.

Этот код получает активный лист, создает две фигуры с разными цветами заливки, группирует их, получает тип класса группы и добавляет текст, указывающий тип класса, к каждой фигуре.

```vba
' VBA Code to replicate OnlyOffice JS functionality

Sub CreateAndGroupShapes()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Define RGB colors
    Dim color1 As Long
    color1 = RGB(255, 111, 61) ' Orange color
    
    Dim color2 As Long
    color2 = RGB(51, 51, 51) ' Dark gray color
    
    ' Add first shape with fill color1
    Dim oShape1 As Shape
    Set oShape1 = oWorksheet.Shapes.AddShape(msoShapeFlowchartNetwork, 60, 35, 200, 100)
    With oShape1
        .Fill.ForeColor.RGB = color1
        .Line.Visible = msoFalse
    End With
    
    ' Add second shape with fill color2
    Dim oShape2 As Shape
    Set oShape2 = oWorksheet.Shapes.AddShape(msoShapeFlowchartNetwork, 60, 35, 200, 100)
    With oShape2
        .Fill.ForeColor.RGB = color2
        .Line.Visible = msoFalse
        .Top = oShape1.Top + 150 ' Position below the first shape
    End With
    
    ' Group the two shapes
    Dim oGroup As ShapeRange
    Set oGroup = oWorksheet.Shapes.Range(Array(oShape1.Name, oShape2.Name)).Group
    
    ' Get class type (using the group name as a placeholder)
    Dim sClassType As String
    sClassType = oGroup.Name
    
    ' Add text to the first shape
    oShape1.TextFrame.Characters.Text = "Class Type = " & sClassType
    
    ' Add text to the second shape
    oShape2.TextFrame.Characters.Text = "Class Type = " & sClassType
End Sub
```

```javascript
// JavaScript Code using OnlyOffice API to create and group shapes

// Get the active worksheet
let oWorksheet = Api.GetActiveSheet();

// Create solid fills with specified RGB colors
let oFill1 = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Orange color
let oFill2 = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)); // Dark gray color

// Create a stroke with no fill
let oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add first shape to the worksheet
let oShape1 = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill1, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Add second shape to the worksheet
let oShape2 = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill2, oStroke, 0, 15 * 36000, 0, 30 * 36000);

// Group the two shapes
let oGroup = oWorksheet.GroupDrawings([oShape1, oShape2]);

// Get the class type of the group
let sClassType = oGroup.GetClassType();

// Add text to the first shape
let oDocContent1 = oShape1.GetDocContent();
let oParagraph1 = oDocContent1.GetElement(0);
oParagraph1.AddText("Class Type = " + sClassType);

// Add text to the second shape
let oDocContent2 = oShape2.GetDocContent();
let oParagraph2 = oDocContent2.GetElement(0);
oParagraph2.AddText("Class Type = " + sClassType);
```