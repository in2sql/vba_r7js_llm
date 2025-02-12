**Description / Описание:**

This code retrieves the active worksheet, creates a gradient fill and stroke, adds a shape to the worksheet, retrieves the class type of the fill, sets column widths, and sets a value in cell A1 indicating the class type.

Этот код получает активный лист, создает градиентную заливку и обводку, добавляет фигуру на лист, получает тип класса заливки, устанавливает ширину столбцов и задает значение в ячейке A1, указывающее тип класса.

```vba
' Excel VBA code equivalent

Sub AddShapeWithGradient()
    Dim oWorksheet As Worksheet
    Dim oGs1 As Object
    Dim oGs2 As Object
    Dim oFill As Object
    Dim oStroke As Object
    Dim sClassType As String
    
    ' Get active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Create gradient stops
    Set oGs1 = CreateGradientStop(RGB(255, 213, 191), 0)
    Set oGs2 = CreateGradientStop(RGB(255, 111, 61), 100000)
    
    ' Create linear gradient fill
    Set oFill = CreateLinearGradientFill(Array(oGs1, oGs2), 5400000)
    
    ' Create stroke with no fill
    Set oStroke = CreateStroke(0, CreateNoFill())
    
    ' Add shape to worksheet
    oWorksheet.Shapes.AddShape(msoShapeFlowchartOfflineStorage, 60 * 36000, 35 * 36000, 100, 100).Fill = oFill
    oWorksheet.Shapes(oWorksheet.Shapes.Count).Line = oStroke
    
    ' Get class type of fill
    sClassType = oFill.ClassType
    
    ' Set column widths
    oWorksheet.Columns(1).ColumnWidth = 15
    oWorksheet.Columns(2).ColumnWidth = 10
    
    ' Set value in cell A1
    oWorksheet.Range("A1").Value = "Class Type = " & sClassType
End Sub

' Function stubs for gradient and stroke creation
Function CreateGradientStop(color As Long, position As Long) As Object
    ' Implement gradient stop creation
End Function

Function CreateLinearGradientFill(stops As Variant, angle As Long) As Object
    ' Implement linear gradient fill creation
End Function

Function CreateStroke(weight As Long, fill As Object) As Object
    ' Implement stroke creation
End Function

Function CreateNoFill() As Object
    ' Implement no fill
End Function
```

```javascript
// OnlyOffice JS code equivalent

// This example gets a class type and inserts it into the document.
var oWorksheet = Api.GetActiveSheet();

// Create gradient stops with specified RGB colors
var oGs1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0);
var oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000);

// Create a linear gradient fill
var oFill = Api.CreateLinearGradientFill([oGs1, oGs2], 5400000);

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet
oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000);

// Get the class type of the fill
var sClassType = oFill.GetClassType();

// Set column widths
oWorksheet.SetColumnWidth(0, 15);
oWorksheet.SetColumnWidth(1, 10);

// Set value in cell A1
oWorksheet.GetRange("A1").SetValue("Class Type = " + sClassType);
```