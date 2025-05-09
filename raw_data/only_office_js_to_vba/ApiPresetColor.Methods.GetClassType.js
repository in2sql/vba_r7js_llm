**Description / Описание**

English: This code retrieves the active sheet, creates a shape with a gradient fill and no stroke, adds it to the sheet, sets column widths, and sets values in cells A1 and B1 to display the class type of a preset color.

Russian: Этот код получает активный лист, создает фигуру с градиентной заливкой и без обводки, добавляет ее на лист, устанавливает ширину столбцов и задает значения в ячейках A1 и B1, чтобы отобразить тип класса предустановленного цвета.

---

```javascript
// This example gets a class type and inserts it into the document.

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a preset color named "peachPuff"
var oPresetColor = Api.CreatePresetColor("peachPuff");

// Create the first gradient stop with the preset color at position 0
var oGs1 = Api.CreateGradientStop(oPresetColor, 0);

// Create the second gradient stop with a custom RGB color at position 100000
var oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000);

// Create a linear gradient fill with the two gradient stops and a rotation of 5400000
var oFill = Api.CreateLinearGradientFill([oGs1, oGs2], 5400000);

// Create a stroke with width 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape of type "flowChartOnlineStorage" to the worksheet with specified parameters
oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000);

// Get the class type of the preset color
var sClassType = oPresetColor.GetClassType();

// Set the width of column A to 15
oWorksheet.SetColumnWidth(0, 15);

// Set the width of column B to 10
oWorksheet.SetColumnWidth(1, 10);

// Set the value of cell A1 to display "Class Type = "
oWorksheet.GetRange("A1").SetValue("Class Type = ");

// Set the value of cell B1 to display the class type of the preset color
oWorksheet.GetRange("B1").SetValue(sClassType);
```

---

```vba
' This example retrieves the active sheet, creates a shape with a gradient fill and no stroke,
' adds it to the sheet, sets column widths, and sets values in cells A1 and B1 to display
' the class type of a preset color.

Sub InsertShapeAndSetValues()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Define the preset color "peachPuff" using RGB
    Dim presetColor As Long
    presetColor = RGB(255, 218, 185) ' peachPuff
    
    ' Add a flowchart shape to the worksheet
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(Type:=msoShapeFlowchartProcess, _
                                           Left:=60, Top:=35, Width:=200, Height:=100)
    
    ' Set gradient fill for the shape
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = presetColor
        .BackColor.RGB = RGB(255, 111, 61) ' Custom RGB color
        .TwoColorGradient Style:=msoGradientHorizontal, Variant:=1
        .GradientStops(1).Position = 0
        .GradientStops(2).Position = 1
    End With
    
    ' Remove the shape's outline (stroke)
    With oShape.Line
        .Visible = msoFalse
    End With
    
    ' Set the width of column A to 15
    oWorksheet.Columns("A").ColumnWidth = 15
    
    ' Set the width of column B to 10
    oWorksheet.Columns("B").ColumnWidth = 10
    
    ' Set the value of cell A1
    oWorksheet.Range("A1").Value = "Class Type = "
    
    ' Since Excel VBA does not have a direct equivalent for GetClassType,
    ' we'll set a placeholder value for demonstration
    oWorksheet.Range("B1").Value = "PresetColorType"
End Sub
```