**Description / Описание:**  
This code changes the position of a drawing object in the active worksheet.  
Этот код изменяет положение объекта рисования на активном листе.

```vba
' VBA code to change the position of a drawing object in the active worksheet
Sub ChangeDrawingPosition()
    Dim oWorksheet As Worksheet
    Dim oShape As Shape
    Dim fillColor As Long
    Dim lineWeight As Single
    Dim lineColor As Long
    
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Define the fill color using RGB
    fillColor = RGB(255, 111, 61)
    
    ' Define the line properties
    lineWeight = 0 ' No line weight
    lineColor = RGB(255, 255, 255) ' No fill for the line
    
    ' Add a flowchart storage shape to the worksheet
    ' msoShapeFlowchartData is used as an example; adjust as needed
    Set oShape = oWorksheet.Shapes.AddShape(Type:=msoShapeFlowchartData, _
                                           Left:=60 * 36000 / 914400 * 72, _ ' Convert EMU to points
                                           Top:=35 * 36000 / 914400 * 72, _ ' Convert EMU to points
                                           Width:=120 * 36000 / 914400 * 72, _ ' Convert EMU to points
                                           Height:=70 * 36000 / 914400 * 72) ' Convert EMU to points
    
    ' Set the fill color of the shape
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = fillColor
        .Solid
    End With
    
    ' Set the line (stroke) properties of the shape
    With oShape.Line
        .Visible = msoFalse
    End With
    
    ' Set the position of the shape
    oShape.Left = 2 * 36000 / 914400 * 72 ' Convert EMU to points
    oShape.Top = 2 * 36000 / 914400 * 72 ' Convert EMU to points
End Sub
```

```javascript
// This example changes the position for the drawing object.
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified parameters
var oDrawing = oWorksheet.AddShape("flowChartOnlineStorage", 
                                   60 * 36000, 
                                   35 * 36000, 
                                   oFill, 
                                   oStroke, 
                                   0, 
                                   2 * 36000, 
                                   0, 
                                   3 * 36000);

// Set the size of the drawing
oDrawing.SetSize(120 * 36000, 70 * 36000);

// Set the position of the drawing
oDrawing.SetPosition(0, 2 * 36000, 2, 3 * 36000);
```