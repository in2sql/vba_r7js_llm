**Description / Описание**

This code creates a linear gradient fill and applies it to a shape in the active worksheet.

Этот код создает заливку с линейным градиентом и применяет ее к фигуре на активном листе.

```vba
' VBA Code to create a linear gradient fill and apply it to a shape in Excel

Sub ApplyLinearGradientFill()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Define gradient colors
    Dim color1 As Long
    Dim color2 As Long
    color1 = RGB(255, 213, 191) ' Light orange
    color2 = RGB(255, 111, 61)  ' Dark orange
    
    ' Add a shape to the worksheet
    Dim shp As Shape
    Set shp = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, _
                                        60, 35, 200, 100) ' Adjust size and position as needed
    
    ' Apply linear gradient fill to the shape
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = color1
        .BackColor.RGB = color2
        .TwoColorGradient msoGradientHorizontal, 1
    End With
    
    ' Remove the stroke (outline) from the shape
    With shp.Line
        .Visible = msoFalse
    End With
End Sub
```

```javascript
// JavaScript Code to create a linear gradient fill and apply it to a shape in OnlyOffice

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Define gradient stops with colors and positions
var oGs1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0);       // Light orange at position 0
var oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000);   // Dark orange at position 100000

// Create a linear gradient fill with the defined gradient stops
var oFill = Api.CreateLinearGradientFill([oGs1, oGs2], 5400000);

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with the gradient fill and no stroke
oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000);
```