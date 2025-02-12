# Description / Описание

**English**: This code retrieves the active worksheet, creates a solid fill and stroke, adds a shape to the worksheet, retrieves the parent sheet of the shape, and adds a text paragraph indicating the name of the parent sheet.

**Русский**: Этот код получает активный лист, создает заливку и обводку, добавляет фигуру на лист, получает родительский лист фигуры и добавляет текстовый абзац с указанием имени родительского листа.

```javascript
// This example shows how to get the drawing's parent sheet.
let oWorksheet = Api.GetActiveSheet(); // Get the active sheet
let oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill with RGB color
let oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
let oShape = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000); // Add a shape to the worksheet
let oParentSheet = oShape.GetParentSheet(); // Get the parent sheet of the shape

let oDocContent = oShape.GetDocContent(); // Get the document content of the shape
let oParagraph = oDocContent.GetElement(0); // Get the first paragraph element
oParagraph.AddText("Parent sheet name is " + oParentSheet.GetName()); // Add text to the paragraph
```

```vba
' This example shows how to get the drawing's parent sheet.
Sub AddShapeAndRetrieveParentSheet()
    ' Get the active sheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Create a solid fill with RGB color
    Dim oFillColor As Long
    oFillColor = RGB(255, 111, 61)
    
    ' Add a shape to the worksheet with specified properties
    ' Parameters: msoShapeFlowchartOfflineStorage, Left, Top, Width, Height
    ' Adjust values as per OnlyOffice example (assuming the units are points)
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartOfflineStorage, 60, 35, 200, 100)
    
    ' Set the fill color
    With oShape.Fill
        .ForeColor.RGB = oFillColor
        .Solid
    End With
    
    ' Set the stroke
    With oShape.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0) ' Assuming CreateNoFill translates to black or no color
        .Weight = 0
    End With
    
    ' Get the parent sheet of the shape
    Dim oParentSheet As Worksheet
    Set oParentSheet = oShape.Parent
    
    ' Add text to the shape indicating the parent sheet's name
    oShape.TextFrame.Characters.Text = "Parent sheet name is " & oParentSheet.Name
End Sub
```