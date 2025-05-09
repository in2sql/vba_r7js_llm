**Description / Описание**

This code demonstrates how to add a shape to the active worksheet, set its size and position, retrieve its width, and display the width in cell A1.
Этот код демонстрирует, как добавить форму на активный лист, установить ее размер и позицию, получить ее ширину и отобразить ширину в ячейке A1.

```vba
' VBA Code to add a shape, set its size and position, retrieve width, and display it in A1

Sub AddShapeAndGetWidth()
    Dim oWorksheet As Worksheet
    Dim oShape As Shape
    Dim nWidth As Double
    Dim oFill As Object
    Dim oStroke As Object

    ' Set the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet

    ' Create a solid fill with RGB color (255, 111, 61)
    Set oFill = CreateSolidFill(255, 111, 61)

    ' Create a stroke with weight 0 and no fill
    Set oStroke = CreateStroke(0, CreateNoFill())

    ' Add a shape to the worksheet
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartProcess, _
        60 * 36, 35 * 36, oFill, oStroke, 0, 2 * 36, 0, 3 * 36)

    ' Set the size of the shape
    oShape.Width = 120 * 36
    oShape.Height = 70 * 36

    ' Set the position of the shape
    oShape.Left = 0
    oShape.Top = 2 * 36
    oShape.Placement = xlMoveAndSize

    ' Get the width of the shape
    nWidth = oShape.Width

    ' Set the value of cell A1 to display the width
    oWorksheet.Range("A1").Value = "Drawing width = " & nWidth
End Sub

' Helper functions to create fill and stroke (placeholders)
Function CreateSolidFill(red As Integer, green As Integer, blue As Integer) As Object
    ' Implement solid fill creation
End Function

Function CreateStroke(weight As Integer, fill As Object) As Object
    ' Implement stroke creation
End Function

Function CreateNoFill() As Object
    ' Implement no fill creation
End Function
```

```javascript
// JavaScript Code to add a shape, set its size and position, retrieve width, and display it in A1

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with weight 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet
var oDrawing = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Set the size of the shape
oDrawing.SetSize(120 * 36000, 70 * 36000);

// Set the position of the shape
oDrawing.SetPosition(0, 2 * 36000, 1, 3 * 36000);

// Get the width of the shape
var nWidth = oDrawing.GetWidth();

// Set the value of cell A1 to display the width
oWorksheet.GetRange("A1").SetValue("Drawing width = " + nWidth);
```