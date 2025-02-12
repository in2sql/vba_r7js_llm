**Description:**
This code creates a complex color scheme by selecting one of the available schemes and adds a 'curvedUpArrow' shape to the active worksheet.
Этот код создает сложную цветовую схему, выбирая одну из доступных схем, и добавляет фигуру 'curvedUpArrow' на активный лист.

```vba
' VBA Code to create a complex color scheme and add a shape to the active worksheet

Sub AddCurvedUpArrowShape()
    Dim oWorksheet As Worksheet
    Dim oSchemeColor As Long
    Dim oFill As Long
    Dim oStroke As Long
    
    ' Set the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Define the color (example using RGB for "dk1")
    oSchemeColor = RGB(0, 0, 139) ' Dark blue
    
    ' Set fill color
    oFill = oSchemeColor
    
    ' Set no stroke
    oStroke = RGB(255, 255, 255) ' White color for no fill simulation
    
    ' Add the 'curvedUpArrow' shape
    oWorksheet.Shapes.AddShape(msoShapeUpArrow, _
        60, 35, _ ' Position (left, top) in points
        200, 100) _  ' Width, Height in points
        .Fill.ForeColor.RGB = oFill
        .Line.Visible = msoFalse
End Sub
```

```javascript
// JavaScript Code to create a complex color scheme and add a 'curvedUpArrow' shape to the active worksheet

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a scheme color with the identifier "dk1"
var oSchemeColor = Api.CreateSchemeColor("dk1");

// Create a solid fill using the scheme color
var oFill = Api.CreateSolidFill(oSchemeColor);

// Create a stroke with width 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a 'curvedUpArrow' shape to the worksheet with specified properties
oWorksheet.AddShape(
    "curvedUpArrow",             // Shape type
    60 * 36000,                  // Left position
    35 * 36000,                  // Top position
    oFill,                       // Fill style
    oStroke,                     // Stroke style
    0,                           // Rotation
    2 * 36000,                   // Width
    1,                           // Height multiplier
    3 * 36000                    // Additional property
);
```