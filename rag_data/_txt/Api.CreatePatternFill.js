**Description / Описание**

English: This code creates a pattern fill and applies it to a flowchart shape in the active worksheet.

Russian: Этот код создает заливку с узором и применяет ее к фигуре блок-схемы на активном листе.

```vba
' VBA equivalent code
' Creates a pattern fill and applies it to a flowchart shape in the active worksheet

Sub AddPatternFillToShape()
    Dim oSheet As Worksheet
    Dim oShape As Shape
    Dim patternForeground As Long
    Dim patternBackground As Long
    
    ' Set the active worksheet
    Set oSheet = ActiveSheet
    
    ' Define the pattern colors
    patternForeground = RGB(255, 111, 61) ' Foreground color (equivalent to CreateRGBColor(255, 111, 61))
    patternBackground = RGB(51, 51, 51)   ' Background color (equivalent to CreateRGBColor(51, 51, 51))
    
    ' Add a flowchart shape to the worksheet
    Set oShape = oSheet.Shapes.AddShape(Type:=msoShapeFlowchartOnlineStorage, _
                                       Left:=60, Top:=35, Width:=150, Height:=100)
    
    ' Apply pattern fill to the shape
    With oShape.Fill
        .Visible = msoTrue
        .Patterned = msoPatternPercent50 ' Equivalent pattern to "dashDnDiag"
        .ForeColor.RGB = patternForeground
        .BackColor.RGB = patternBackground
    End With
    
    ' Apply no fill to the stroke (border) of the shape
    With oShape.Line
        .Visible = msoTrue
        .Weight = 0 ' Equivalent to CreateStroke with weight 0
        .ForeColor.RGB = RGB(0, 0, 0) ' Default border color
        .Transparency = 1 ' Makes the stroke invisible, equivalent to CreateNoFill
    End With
End Sub
```

```javascript
// JavaScript code using OnlyOffice API

// This example creates a pattern fill to apply to the object using the selected pattern as the object background.
// Этот пример создает заливку с узором и применяет ее к объекту, используя выбранный узор в качестве фона объекта.

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreatePatternFill("dashDnDiag", Api.CreateRGBColor(255, 111, 61), Api.CreateRGBColor(51, 51, 51)); // Create a pattern fill with specified colors
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with weight 0 and no fill
oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000); // Add the shape to the worksheet with the fill and stroke
```