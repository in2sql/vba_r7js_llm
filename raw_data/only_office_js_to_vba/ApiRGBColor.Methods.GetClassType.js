# Script that gets a class type, inserts a shape with gradient fill, sets column widths, and writes values to cells A1 and B1 in the worksheet.
# Скрипт, который получает тип класса, вставляет фигуру с градиентной заливкой, устанавливает ширину столбцов и записывает значения в ячейки A1 и B1 рабочего листа.

```vba
' VBA Code
Sub InsertShapeAndSetValues()
    Dim oWorksheet As Worksheet
    Dim oRGBColor As Long
    Dim oFill As Object
    Dim oStroke As Object
    Dim sClassType As String
    
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Create RGB color
    oRGBColor = RGB(255, 213, 191)
    
    ' Create gradient fill
    ' Excel VBA does not have a direct gradient fill creation method, so we use built-in gradient
    With oWorksheet.Shapes.AddShape(msoShapeFlowchartProcess, 60, 35, 150, 75)
        ' Set fill to gradient
        .Fill.TwoColorGradient msoGradientHorizontal, 1
        .Fill.ForeColor.RGB = RGB(255, 213, 191)
        .Fill.BackColor.RGB = RGB(255, 111, 61)
        ' Set no line
        .Line.Visible = msoFalse
    End With
    
    ' Assuming class type is a string representation of the color
    sClassType = "RGB(255, 213, 191)"
    
    ' Set column widths
    oWorksheet.Columns(1).ColumnWidth = 15
    oWorksheet.Columns(2).ColumnWidth = 10
    
    ' Set cell values
    oWorksheet.Range("A1").Value = "Class Type = "
    oWorksheet.Range("B1").Value = sClassType
End Sub
```

```javascript
// OnlyOffice JS Code
// This example gets a class type and inserts it into the document.

var oWorksheet = Api.GetActiveSheet();

// Create an RGB color
var oRGBColor = Api.CreateRGBColor(255, 213, 191);

// Create gradient stops
var oGs1 = Api.CreateGradientStop(oRGBColor, 0);
var oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000);

// Create a linear gradient fill
var oFill = Api.CreateLinearGradientFill([oGs1, oGs2], 5400000);

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape with the specified properties
oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000);

// Get class type from RGB color
var sClassType = oRGBColor.GetClassType();

// Set column widths
oWorksheet.SetColumnWidth(0, 15);
oWorksheet.SetColumnWidth(1, 10);

// Set cell values
oWorksheet.GetRange("A1").SetValue("Class Type = ");
oWorksheet.GetRange("B1").SetValue(sClassType);
```