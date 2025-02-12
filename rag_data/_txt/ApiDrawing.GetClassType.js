**Description / Описание:**  
This code retrieves the active worksheet, creates a solid fill color, defines a stroke, adds a shape with specific properties, sets its size and position, retrieves its class type, adjusts column widths, and sets a cell value with the class type.  
Этот код получает активный лист, создает сплошной цвет заливки, определяет обводку, добавляет фигуру с определенными свойствами, устанавливает ее размер и положение, получает тип класса, настраивает ширину столбцов и устанавливает значение ячейки с типом класса.

---

```vba
' VBA Code Equivalent

Sub InsertShapeAndSetClassType()
    Dim oWorksheet As Worksheet
    Dim oShape As Shape
    Dim sClassType As String
    
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Add a shape to the worksheet
    ' msoShapeFlowchartOnlineStorage corresponds to "flowChartOnlineStorage"
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, _
                                           60, 35, _
                                           120, 70)
    
    ' Set the fill color to RGB(255, 111, 61)
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61)
        .Solid
    End With
    
    ' Set the line (stroke) to no fill
    With oShape.Line
        .Visible = msoFalse
    End With
    
    ' Set the size of the shape
    oShape.Width = 120
    oShape.Height = 70
    
    ' Set the position of the shape
    oShape.Left = 0
    oShape.Top = 72 ' 2 * 36 (assuming 36 points per 36000 units)
    
    ' Retrieve the class type (using the ShapeType as an equivalent)
    sClassType = TypeName(oShape)
    
    ' Set column widths
    oWorksheet.Columns(1).ColumnWidth = 15
    oWorksheet.Columns(2).ColumnWidth = 10
    
    ' Set the value of cell A1
    oWorksheet.Range("A1").Value = "Class Type = " & sClassType
End Sub
```

---

```javascript
// This example gets a class type and inserts it into the document.
// Получает тип класса и вставляет его в документ.

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
var oDrawing = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000); // Add a shape to the worksheet
oDrawing.SetSize(120 * 36000, 70 * 36000); // Set the size of the shape
oDrawing.SetPosition(0, 2 * 36000, 1, 3 * 36000); // Set the position of the shape
var sClassType = oDrawing.GetClassType(); // Get the class type of the shape
oWorksheet.SetColumnWidth(0, 15); // Set the width of the first column
oWorksheet.SetColumnWidth(1, 10); // Set the width of the second column
oWorksheet.GetRange("A1").SetValue("Class Type = " + sClassType); // Set the value of cell A1
```