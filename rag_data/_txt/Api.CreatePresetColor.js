**Description (English):**
This code creates a shape with a linear gradient fill by selecting preset and RGB colors, and adds it to the active worksheet.

**Описание (Russian):**
Этот код создает фигуру с линейной градиентной заливкой, выбирая предустановленные и RGB цвета, и добавляет ее на активный лист.

```javascript
// JavaScript code for OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a preset color "peachPuff"
var oPresetColor = Api.CreatePresetColor("peachPuff");

// Create the first gradient stop with the preset color at position 0
var oGs1 = Api.CreateGradientStop(oPresetColor, 0);

// Create the second gradient stop with an RGB color at position 100%
var oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000);

// Create a linear gradient fill with the two gradient stops and a gradient style
var oFill = Api.CreateLinearGradientFill([oGs1, oGs2], 5400000);

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified parameters and the created fill and stroke
oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000);
```

```vba
' VBA equivalent code

' Get the active worksheet
Dim oWorksheet As Object
Set oWorksheet = Api.GetActiveSheet()

' Create a preset color "peachPuff"
Dim oPresetColor As Object
Set oPresetColor = Api.CreatePresetColor("peachPuff")

' Create the first gradient stop with the preset color at position 0
Dim oGs1 As Object
Set oGs1 = Api.CreateGradientStop(oPresetColor, 0)

' Create the second gradient stop with an RGB color at position 100%
Dim oGs2 As Object
Set oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000)

' Create a linear gradient fill with the two gradient stops and a gradient style
Dim oFill As Object
Set oFill = Api.CreateLinearGradientFill(Array(oGs1, oGs2), 5400000)

' Create a stroke with no fill
Dim oStroke As Object
Set oStroke = Api.CreateStroke(0, Api.CreateNoFill())

' Add a shape to the worksheet with specified parameters and the created fill and stroke
oWorksheet.AddShape "flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000
```