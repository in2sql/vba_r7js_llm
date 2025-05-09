**Description:**
This VBA code performs the same operations as the provided OnlyOffice JavaScript code: it creates a shape with specified fill and stroke properties, sets its size and position, retrieves its class type, sets column widths, and writes the class type to cell A1.

Этот VBA код выполняет те же операции, что и предоставленный JavaScript код OnlyOffice: он создает фигуру с заданными свойствами заливки и обводки, устанавливает её размер и позицию, получает тип класса, устанавливает ширину столбцов и записывает тип класса в ячейку A1.

```vba
' VBA Code Equivalent to OnlyOffice JS Example

Sub InsertShapeAndSetProperties()
    Dim oWorksheet As Worksheet
    Dim oShape As Shape
    Dim sClassType As String
    
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Add a shape to the worksheet
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartDatabase, _
        Range("C2").Left, Range("C2").Top, 120, 70)
    
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
    
    ' Set the size of the shape
    oShape.Width = 120
    oShape.Height = 70
    
    ' Set the position of the shape
    oShape.Left = Range("A2").Left
    oShape.Top = Range("A2").Top
    
    ' Get the class type (using the AutoShape Type as a placeholder)
    sClassType = oShape.AutoShapeType
    
    ' Set column widths
    oWorksheet.Columns(1).ColumnWidth = 15
    oWorksheet.Columns(2).ColumnWidth = 10
    
    ' Set the value of cell A1
    oWorksheet.Range("A1").Value = "Class Type = " & sClassType
End Sub
```

```javascript
// OnlyOffice JS Code Equivalent

// This example gets a class type and inserts it into the document.
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with width 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified properties
var oDrawing = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Set the size of the shape
oDrawing.SetSize(120 * 36000, 70 * 36000);

// Set the position of the shape
oDrawing.SetPosition(0, 2 * 36000, 1, 3 * 36000);

// Get the class type of the shape
var sClassType = oDrawing.GetClassType();

// Set column widths
oWorksheet.SetColumnWidth(0, 15);
oWorksheet.SetColumnWidth(1, 10);

// Set the value of cell A1 with the class type
oWorksheet.GetRange("A1").SetValue("Class Type = " + sClassType);
```