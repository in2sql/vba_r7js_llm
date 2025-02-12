**Description:  
This code creates a gradient-filled shape on the active worksheet with specified gradient stops and stroke properties.  
Этот код создает фигуру с градиентной заливкой на активном листе, используя заданные градиентные остановки и свойства обводки.**

```vba
' VBA Code to create a gradient-filled shape on the active worksheet

Sub AddGradientShape()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Define gradient colors
    Dim color1 As Long
    Dim color2 As Long
    color1 = RGB(255, 213, 191) ' Light color
    color2 = RGB(255, 111, 61)  ' Dark color
    
    ' Add a flowchart storage shape
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartStorage, _
                                 60, 35, 360, 300) ' Left, Top, Width, Height
    
    ' Set gradient fill
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = color1
        .BackColor.RGB = color2
        .TwoColorGradient msoGradientHorizontal, 1
        .Angle = 540 ' Adjust angle as needed
    End With
    
    ' Set no fill for the stroke
    With shp.Line
        .Weight = 0
        .Visible = msoFalse
    End With
End Sub
```

```javascript
// JavaScript Code to create a gradient-filled shape on the active worksheet

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create gradient stops with specified colors and positions
var oGs1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0); // Light color at position 0
var oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000); // Dark color at position 100000

// Create a linear gradient fill with the gradient stops and angle
var oFill = Api.CreateLinearGradientFill([oGs1, oGs2], 5400000); // Angle specified as 5400000

// Create a stroke with weight 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add the shape to the worksheet with specified properties
oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000); 
```