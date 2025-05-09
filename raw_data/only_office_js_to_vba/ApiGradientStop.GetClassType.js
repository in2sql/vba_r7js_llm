# Description / Описание

**English:**  
This code adds a shaped object with a gradient fill to the active worksheet, sets column widths, and places a value in cell A1 indicating the class type of the gradient stop object.

**Русский:**  
Этот код добавляет объект с градиентной заливкой на активный лист, устанавливает ширину столбцов и помещает значение в ячейку A1, указывающее тип класса объекта градиентной остановки.

```vba
' VBA Code Equivalent

Sub AddShapeWithGradientAndSetValues()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Define colors using RGB
    Dim color1 As Long
    Dim color2 As Long
    color1 = RGB(255, 213, 191) ' Light Orange
    color2 = RGB(255, 111, 61)  ' Dark Orange
    
    ' Add a shape to the worksheet
    Dim shp As Shape
    Set shp = oWorksheet.Shapes.AddShape(msoShapeFlowchartStorage, _
        60 * 36, 35 * 36, 200, 100) ' Positions converted from OnlyOffice units to points
    
    ' Apply linear gradient fill
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = color1
        .TwoColorGradient msoGradientHorizontal, 1
        .GradientStops.Insert RGB(255, 111, 61), 1 ' Adding second gradient color
    End With
    
    ' Remove the shape's stroke
    With shp.Line
        .Visible = msoFalse
    End With
    
    ' Set column widths
    oWorksheet.Columns(1).ColumnWidth = 15
    oWorksheet.Columns(2).ColumnWidth = 10
    
    ' Get the class type (Using TypeName as VBA equivalent)
    Dim sClassType As String
    sClassType = TypeName(color1)
    
    ' Set value in cell A1
    oWorksheet.Range("A1").Value = "Class Type = " & sClassType
End Sub
```

```javascript
// JavaScript Code Equivalent using OnlyOffice API

// This function adds a shaped object with a gradient fill, sets column widths, and sets a cell value with the class type.
function addShapeAndSetValues() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Create gradient stops with specific colors and positions
    var oGs1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0);
    var oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000);
    
    // Create a linear gradient fill with the gradient stops
    var oFill = Api.CreateLinearGradientFill([oGs1, oGs2], 5400000);
    
    // Create a stroke with no fill
    var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    
    // Add a shape to the worksheet with the specified properties
    oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000);
    
    // Get the class type of the first gradient stop
    var sClassType = oGs1.GetClassType();
    
    // Set the width of the first two columns
    oWorksheet.SetColumnWidth(0, 15);
    oWorksheet.SetColumnWidth(1, 10);
    
    // Set the value of cell A1 to display the class type
    oWorksheet.GetRange("A1").SetValue("Class Type = " + sClassType); 
}

// Call the function to execute the actions
addShapeAndSetValues();
```