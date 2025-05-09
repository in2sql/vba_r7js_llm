**Description:**
This script modifies the position and size of a drawing object in the active worksheet.
Этот скрипт изменяет позицию и размер объекта рисования в активном листе.

```vba
' VBA Code to modify the position and size of a drawing object in Excel

Sub ModifyDrawingObject()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Create a shape with specific properties
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, _
        60 * 36000, 35 * 36000, 120 * 36000, 70 * 36000)
    
    ' Set the fill color using RGB
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61)
        .Solid
    End With
    
    ' Set the line (stroke) properties
    With oShape.Line
        .Visible = msoTrue
        .Weight = 0
        .ForeColor.RGB = RGB(255, 255, 255) ' No fill equivalent
    End With
    
    ' Set the position of the shape
    oShape.Left = 2 * 36000
    oShape.Top = 3 * 36000
End Sub
```

```javascript
// JavaScript Code to modify the position and size of a drawing object in OnlyOffice

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified properties
var oDrawing = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Set the size of the drawing
oDrawing.SetSize(120 * 36000, 70 * 36000);

// Set the position of the drawing
oDrawing.SetPosition(0, 2 * 36000, 2, 3 * 36000);
```