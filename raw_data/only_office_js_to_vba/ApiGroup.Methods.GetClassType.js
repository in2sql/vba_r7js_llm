**Description / Описание**

*English:* This script creates two shapes with different fill colors on the active sheet, groups them, retrieves the class type of the group, and adds text displaying the class type to each shape.

*Russian:* Этот скрипт создает две фигуры с разными цветами заливки на активном листе, группирует их, получает тип класса группы и добавляет текст, отображающий тип класса, к каждой фигуре.

```vba
' VBA Code equivalent to the provided OnlyOffice JS example

Sub CreateAndGroupShapes()
    ' Get the active sheet
    Dim oWorksheet As Object
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Create solid fill colors
    Dim oFill1 As Object
    Set oFill1 = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61))
    
    Dim oFill2 As Object
    Set oFill2 = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51))
    
    ' Create stroke with no fill
    Dim oStroke As Object
    Set oStroke = Api.CreateStroke(0, Api.CreateNoFill())
    
    ' Add first shape to the worksheet
    Dim oShape1 As Object
    Set oShape1 = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill1, oStroke, 0, 2 * 36000, 0, 3 * 36000)
    
    ' Add second shape to the worksheet
    Dim oShape2 As Object
    Set oShape2 = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill2, oStroke, 0, 15 * 36000, 0, 30 * 36000)
    
    ' Group the two shapes
    Dim oGroup As Object
    Set oGroup = oWorksheet.GroupDrawings(Array(oShape1, oShape2))
    
    ' Get class type of the group
    Dim sClassType As String
    sClassType = oGroup.GetClassType()
    
    ' Add text to the first shape
    Dim oDocContent1 As Object
    Set oDocContent1 = oShape1.GetDocContent()
    
    Dim oParagraph1 As Object
    Set oParagraph1 = oDocContent1.GetElement(0)
    
    oParagraph1.AddText "Class Type = " & sClassType
    
    ' Add text to the second shape
    Dim oDocContent2 As Object
    Set oDocContent2 = oShape2.GetDocContent()
    
    Dim oParagraph2 As Object
    Set oParagraph2 = oDocContent2.GetElement(0)
    
    oParagraph2.AddText "Class Type = " & sClassType
End Sub
```

```javascript
// OnlyOffice JS Code equivalent to the provided example

// This script creates two shapes with different fill colors on the active sheet, groups them, retrieves the class type of the group, and adds text displaying the class type to each shape.

let oWorksheet = Api.GetActiveSheet();

// Create solid fill colors
let oFill1 = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
let oFill2 = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));

// Create stroke with no fill
let oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add first shape to the worksheet
let oShape1 = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill1, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Add second shape to the worksheet
let oShape2 = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill2, oStroke, 0, 15 * 36000, 0, 30 * 36000);

// Group the two shapes
let oGroup = oWorksheet.GroupDrawings([oShape1, oShape2]);

// Get class type of the group
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