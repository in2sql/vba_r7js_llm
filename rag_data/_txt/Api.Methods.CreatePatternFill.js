# Description

This code creates a pattern fill and adds a flowchart shape to the active worksheet.
Этот код создает заливку паттерном и добавляет элемент блок-схемы на активный лист.

```vba
' VBA Code to create a pattern fill and add a shape to the active worksheet

Sub AddPatternFillShape()
    Dim oWorksheet As Worksheet
    Dim oShape As Shape
    
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Add a flowchart online storage shape to the worksheet with specified dimensions
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, 60, 35, 200, 100)
    
    ' Set the pattern fill with foreground and background colors
    With oShape.Fill
        .Pattern = msoPatternDashDownDiagonal
        .ForeColor.RGB = RGB(255, 111, 61)
        .BackColor.RGB = RGB(51, 51, 51)
    End With
    
    ' Set the line properties with no fill and weight of 0
    With oShape.Line
        .Weight = 0
        .Visible = msoFalse
    End With
End Sub
```

```javascript
// This example creates a pattern fill to apply to the object using the selected pattern as the object background.
var oWorksheet = Api.GetActiveSheet();
// Create a pattern fill with specified pattern and colors
var oFill = Api.CreatePatternFill("dashDnDiag", Api.CreateRGBColor(255, 111, 61), Api.CreateRGBColor(51, 51, 51));
// Create a stroke with no fill and weight of 0
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
// Add a flowchart online storage shape with specified dimensions and formatting
oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000);
```