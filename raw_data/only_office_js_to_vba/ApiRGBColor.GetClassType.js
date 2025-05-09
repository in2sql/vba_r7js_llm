### Description / Описание
This code retrieves the active worksheet, creates a gradient-filled shape, sets column widths, and updates specific cell values with the class type of a color object.
Этот код получает активный лист, создает фигуру с градиентной заливкой, устанавливает ширину столбцов и обновляет значения определенных ячеек типом класса объекта цвета.

```javascript
// This example gets a class type and inserts it into the document.
// Этот пример получает тип класса и вставляет его в документ.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oRGBColor = Api.CreateRGBColor(255, 213, 191); // Create an RGB color
var oGs1 = Api.CreateGradientStop(oRGBColor, 0); // Create first gradient stop
var oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000); // Create second gradient stop
var oFill = Api.CreateLinearGradientFill([oGs1, oGs2], 5400000); // Create linear gradient fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create stroke with no fill
oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000); // Add shape to worksheet
var sClassType = oRGBColor.GetClassType(); // Get the class type of the color
oWorksheet.SetColumnWidth(0, 15); // Set width of column A
oWorksheet.SetColumnWidth(1, 10); // Set width of column B
oWorksheet.GetRange("A1").SetValue("Class Type = "); // Set value in cell A1
oWorksheet.GetRange("B1").SetValue(sClassType); // Set value in cell B1
```

```vba
' This code retrieves the active worksheet, creates a gradient-filled shape,
' sets column widths, and updates specific cell values with the class type of a color object.
' Этот код получает активный лист, создает фигуру с градиентной заливкой,
' устанавливает ширину столбцов и обновляет значения определенных ячеек типом класса объекта цвета.

Sub InsertShapeAndSetValues()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet ' Get the active worksheet
    
    ' Define colors
    Dim color1 As Long
    color1 = RGB(255, 213, 191) ' Create RGB color
    
    Dim color2 As Long
    color2 = RGB(255, 111, 61) ' Create second RGB color
    
    ' Add a shape with gradient fill
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartOnlineStorage, 60 * 36, 35 * 36, 200, 100) ' Add shape to worksheet
    With shp.Fill
        .TwoColorGradient msoGradientHorizontal, 1 ' Create linear gradient fill
        .ForeColor.RGB = color1
        .BackColor.RGB = color2
    End With
    shp.Line.Visible = msoFalse ' Create stroke with no fill
    
    ' Get the class type (VBA does not have a direct equivalent, using type name)
    Dim classType As String
    classType = TypeName(color1) ' Get the type name of the color variable
    
    ' Set column widths
    ws.Columns(1).ColumnWidth = 15 ' Set width of column A
    ws.Columns(2).ColumnWidth = 10 ' Set width of column B
    
    ' Set cell values
    ws.Range("A1").Value = "Class Type = " ' Set value in cell A1
    ws.Range("B1").Value = classType ' Set value in cell B1
End Sub
```