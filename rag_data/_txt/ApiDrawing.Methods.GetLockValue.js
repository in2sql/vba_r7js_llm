**Description / Описание**

English: This example demonstrates how to get and set the lock value for a specified lock type of a drawing in a worksheet.

Russian: Этот пример демонстрирует, как получить и установить значение блокировки для указанного типа блокировки рисунка на листе.

```vba
' VBA Code to get and set the lock value for a specified lock type of a drawing

Sub ManageDrawingLock()
    Dim oWorksheet As Worksheet
    Dim oShape As Shape
    Dim lockValue As Boolean
    
    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Add a shape to the worksheet
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartProcess, _
        Left:=60, Top:=35, Width:=120, Height:=70)
    
    ' Set the fill color using RGB
    oShape.Fill.ForeColor.RGB = RGB(255, 111, 61)
    
    ' Remove the stroke
    oShape.Line.Visible = msoFalse
    
    ' Set the lock property to prevent selection
    oShape.Locked = True
    oShape.LockAspectRatio = msoFalse
    oShape.Selectable = False
    
    ' Get the lock value
    lockValue = Not oShape.Locked ' Assuming LockAspectRatio as lock type
    
    ' Set the value in cell A1
    oWorksheet.Range("A1").Value = "This drawing cannot be selected: " & lockValue
End Sub
```

```javascript
// JS Code to get and set the lock value for a specified lock type of a drawing

// This example shows how to get the lock value for the specified lock type of the drawing.
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet
var oDrawing = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Set the size of the shape
oDrawing.SetSize(120 * 36000, 70 * 36000);

// Set the position of the shape
oDrawing.SetPosition(0, 2 * 36000, 1, 3 * 36000);

// Set the lock value to prevent selection
oDrawing.SetLockValue("noSelect", true);

// Get the lock value
var bLockValue = oDrawing.GetLockValue("noSelect");

// Set the value in cell A1 with the lock status
oWorksheet.GetRange("A1").SetValue("This drawing cannot be selected: " + bLockValue);
```