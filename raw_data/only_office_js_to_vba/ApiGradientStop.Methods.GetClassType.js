# Description / Описание
This code retrieves the active worksheet, creates a gradient fill, adds a flowchart shape with the specified fill and stroke, sets column widths, and writes the class type to cell A1.
Этот код получает активный лист, создает градиентную заливку, добавляет форму блок-схемы с указанной заливкой и обводкой, устанавливает ширину столбцов и записывает тип класса в ячейку A1.

```javascript
// This example gets a class type and inserts it into the document.

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create gradient stops with specified RGB colors and positions
var oGs1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0);
var oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000);

// Create a linear gradient fill with the gradient stops and angle
var oFill = Api.CreateLinearGradientFill([oGs1, oGs2], 5400000);

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a flowchart shape to the worksheet with specified parameters
oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000);

// Get the class type from the first gradient stop
var sClassType = oGs1.GetClassType();

// Set the width of column A to 15 and column B to 10
oWorksheet.SetColumnWidth(0, 15);
oWorksheet.SetColumnWidth(1, 10);

// Set the value of cell A1 to display the class type
oWorksheet.GetRange("A1").SetValue("Class Type = " + sClassType); 
```

```vba
' This VBA code retrieves the active worksheet, creates a gradient fill, adds a flowchart shape, sets column widths, and writes the class type to cell A1.
' Этот VBA код получает активный лист, создает градиентную заливку, добавляет форму блок-схемы, устанавливает ширину столбцов и записывает тип класса в ячейку A1.

Sub AddShapeAndSetValues()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet ' Get the active worksheet
    
    ' Define gradient colors
    Dim color1 As Long
    Dim color2 As Long
    color1 = RGB(255, 213, 191) ' Light color
    color2 = RGB(255, 111, 61)  ' Dark color
    
    ' Add a shape (Flowchart) with gradient fill
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartPredefinedProcess, _
        Left:=60 * 72, Top:=35 * 72, Width:=200, Height:=100) ' Positions are in points (1 inch = 72 points)
    
    ' Create gradient fill
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = color1
        .TwoColorGradient msoGradientHorizontal, 1
        .GradientStops.Clear
        .GradientStops.Insert RGB(255, 213, 191), 0
        .GradientStops.Insert RGB(255, 111, 61), 1
        .Rotation = 60
    End With
    
    ' Remove the stroke
    With shp.Line
        .Visible = msoFalse
    End With
    
    ' Set column widths
    ws.Columns(1).ColumnWidth = 15 ' Column A
    ws.Columns(2).ColumnWidth = 10 ' Column B
    
    ' Get class type (simulated as "GradientStop" since VBA doesn't have GetClassType for gradient stops)
    Dim sClassType As String
    sClassType = "GradientStop"
    
    ' Set value in cell A1
    ws.Range("A1").Value = "Class Type = " & sClassType
End Sub
```