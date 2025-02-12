**Description / Описание:**
This code sets the lock value to the specified lock type of the current drawing.
Этот код устанавливает значение блокировки для указанного типа блокировки текущего рисунка.

```javascript
// Description: This code sets the lock value to the specified lock type of the current drawing.
// Описание: Этот код устанавливает значение блокировки для указанного типа блокировки текущего рисунка.

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill color with RGB values (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified parameters
var oDrawing = oWorksheet.AddShape(
    "flowChartOnlineStorage",
    60 * 36000,
    35 * 36000,
    oFill,
    oStroke,
    0,
    2 * 36000,
    0,
    3 * 36000
);

// Set the size of the drawing
oDrawing.SetSize(120 * 36000, 70 * 36000);

// Set the position of the drawing
oDrawing.SetPosition(0, 2 * 36000, 1, 3 * 36000);

// Lock the drawing to prevent selection
oDrawing.SetLockValue("noSelect", true);

// Retrieve the lock value
var bLockValue = oDrawing.GetLockValue("noSelect");

// Set the value of cell A1 with a message including the lock status
oWorksheet.GetRange("A1").SetValue("This drawing cannot be selected: " + bLockValue);
```

```vba
' Description: This code sets the lock value to the specified lock type of the current drawing.
' Описание: Этот код устанавливает значение блокировки для указанного типа блокировки текущего рисунка.

Sub SetDrawingLock()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Create a solid fill color with RGB values (255, 111, 61)
    Dim oFill As Object
    Set oFill = oWorksheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 100, 100).Fill
    oFill.ForeColor.RGB = RGB(255, 111, 61)
    oFill.Solid
    
    ' Create a stroke with no fill
    Dim oStroke As Object
    Set oStroke = oWorksheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 100, 100).Line
    oStroke.Visible = msoFalse
    
    ' Add a shape to the worksheet with specified parameters
    Dim oDrawing As Shape
    Set oDrawing = oWorksheet.Shapes.AddShape( _
        Type:=msoShapeFlowchartOfflineStorage, _
        Left:=60 * 3.6, _ ' Adjusted for Excel's measurement units
        Top:=35 * 3.6, _
        Width:=120 * 3.6, _
        Height:=70 * 3.6)
    
    ' Apply fill and stroke to the drawing
    With oDrawing
        .Fill.ForeColor.RGB = RGB(255, 111, 61)
        .Fill.Solid
        .Line.Visible = msoFalse
    End With
    
    ' Lock the drawing to prevent selection
    oDrawing.Locked = True
    oDrawing.LockAspectRatio = msoTrue
    oDrawing.Placement = xlMoveAndSize
    
    ' Retrieve the lock value
    Dim bLockValue As Boolean
    bLockValue = oDrawing.Locked
    
    ' Set the value of cell A1 with a message including the lock status
    oWorksheet.Range("A1").Value = "This drawing cannot be selected: " & bLockValue
End Sub
```