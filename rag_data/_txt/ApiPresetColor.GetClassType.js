## Description / Описание

**English:**  
This script retrieves the active worksheet, creates a preset color, defines a linear gradient fill, adds a specific shape to the worksheet, and sets the widths of columns A and B along with assigning values to cells A1 and B1.

**Russian:**  
Этот скрипт получает активный рабочий лист, создает предустановленный цвет, определяет линейный градиентный залив, добавляет определенную фигуру на лист и устанавливает ширины столбцов A и B, а также присваивает значения ячейкам A1 и B1.

### VBA Code

```vba
Sub AddShapeAndSetValues()
    ' Get the active worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Create a preset color (RGB equivalent of "peach puff")
    Dim presetColor As Long
    presetColor = RGB(255, 218, 185)
    
    ' Define gradient colors
    Dim gradientColor1 As Long
    gradientColor1 = presetColor
    
    Dim gradientColor2 As Long
    gradientColor2 = RGB(255, 111, 61)
    
    ' Add a shape with a linear gradient fill
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartOnlineStorage, _
                                 60 * 36, 35 * 36, _
                                 200, 100) ' Width and Height in points
                                 
    ' Define the gradient fill
    With shp.Fill
        .TwoColorGradient msoGradientHorizontal, 1
        .ForeColor.RGB = gradientColor1
        .BackColor.RGB = gradientColor2
        .GradientAngle = 54 ' Equivalent to 5400000 in OnlyOffice (scaled down)
    End With
    
    ' Remove the stroke
    shp.Line.Visible = msoFalse
    
    ' Set column widths
    ws.Columns("A").ColumnWidth = 15
    ws.Columns("B").ColumnWidth = 10
    
    ' Set cell values
    ws.Range("A1").Value = "Class Type = "
    ws.Range("B1").Value = "CustomClassType" ' Replace with actual class type if available
End Sub
```

### OnlyOffice JS Code

```javascript
// This script retrieves the active worksheet, creates a preset color, defines a linear gradient fill,
// adds a specific shape to the worksheet, and sets the widths of columns A and B along with assigning values to cells A1 and B1.

function addShapeAndSetValues() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Create a preset color
    var oPresetColor = Api.CreatePresetColor("peachPuff");
    
    // Create gradient stops
    var oGs1 = Api.CreateGradientStop(oPresetColor, 0);
    var oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000);
    
    // Create a linear gradient fill
    var oFill = Api.CreateLinearGradientFill([oGs1, oGs2], 5400000);
    
    // Create a stroke with no fill
    var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    
    // Add a shape to the worksheet
    oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000);
    
    // Get the class type from the preset color
    var sClassType = oPresetColor.GetClassType();
    
    // Set column widths
    oWorksheet.SetColumnWidth(0, 15);
    oWorksheet.SetColumnWidth(1, 10);
    
    // Set cell values
    oWorksheet.GetRange("A1").SetValue("Class Type = ");
    oWorksheet.GetRange("B1").SetValue(sClassType);
}

// Execute the function
addShapeAndSetValues();
```