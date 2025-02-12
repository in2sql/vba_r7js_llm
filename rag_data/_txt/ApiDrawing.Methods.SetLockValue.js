# Description
**English:** This code creates a shape in the active worksheet with specified fill and stroke properties, sets its size and position, locks it from being selected, and updates cell A1 with a message indicating whether the shape is locked from selection.

**Russian:** Этот код создает фигуру на активном листе с указанными свойствами заливки и обводки, устанавливает ее размер и позицию, блокирует возможность выбора фигуры и обновляет ячейку A1 сообщением, указывающим, заблокирована ли фигура от выбора.

```javascript
// JavaScript code using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a flowchart shape to the worksheet with specified dimensions and styles
var oDrawing = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Set the size of the shape
oDrawing.SetSize(120 * 36000, 70 * 36000);

// Set the position of the shape
oDrawing.SetPosition(0, 2 * 36000, 1, 3 * 36000);

// Lock the shape from being selected
oDrawing.SetLockValue("noSelect", true);

// Get the lock status of the shape
var bLockValue = oDrawing.GetLockValue("noSelect");

// Set the value of cell A1 to indicate the lock status
oWorksheet.GetRange("A1").SetValue("This drawing cannot be selected: " + bLockValue); 
```

```vba
' VBA code equivalent to the OnlyOffice JavaScript example

Sub CreateAndLockShape()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Define RGB color for the fill
    Dim fillColor As Long
    fillColor = RGB(255, 111, 61)
    
    ' Add an oval shape to the worksheet with specified dimensions
    Dim oDrawing As Shape
    Set oDrawing = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, _
                60 * 72, 35 * 72, 120 * 72, 70 * 72) ' Excel uses points; 1 inch = 72 points
    
    ' Set the fill color of the shape
    With oDrawing.Fill
        .Visible = msoTrue
        .ForeColor.RGB = fillColor
        .Solid
    End With
    
    ' Remove the stroke (outline) of the shape
    With oDrawing.Line
        .Visible = msoFalse
    End With
    
    ' Set the position of the shape
    oDrawing.Left = 0 * 72 ' Left position in points
    oDrawing.Top = 2 * 72 ' Top position in points
    
    ' Lock the shape from being selected
    oDrawing.Locked = msoTrue
    oDrawing.Placement = xlMoveAndSize
    
    ' Update cell A1 with lock status
    oWorksheet.Range("A1").Value = "This drawing cannot be selected: " & oDrawing.Locked
End Sub
```