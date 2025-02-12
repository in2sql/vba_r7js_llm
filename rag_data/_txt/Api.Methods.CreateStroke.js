### Description / Описание

**English:** This example creates a stroke by adding shadows to the element.

**Russian:** Этот пример создает обводку, добавляя тени к элементу.

```vba
' VBA code
' This example creates a stroke by adding shadows to the element.

Sub AddShapeWithStroke()
    Dim oWorksheet As Object
    Dim oGs1 As Object
    Dim oGs2 As Object
    Dim oFill As Object
    Dim oFill1 As Object
    Dim oStroke As Object
    
    ' Get the active sheet
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Create gradient stops
    Set oGs1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0)
    Set oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000)
    
    ' Create linear gradient fill
    Set oFill = Api.CreateLinearGradientFill(Array(oGs1, oGs2), 5400000)
    
    ' Create solid fill
    Set oFill1 = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51))
    
    ' Create stroke
    Set oStroke = Api.CreateStroke(3 * 36000, oFill1)
    
    ' Add shape to the worksheet
    oWorksheet.AddShape "flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000
End Sub
```

```javascript
// JavaScript code
// This example creates a stroke by adding shadows to the element.

function addShapeWithStroke() {
    var oWorksheet = Api.GetActiveSheet();
    var oGs1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0);
    var oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000);
    var oFill = Api.CreateLinearGradientFill([oGs1, oGs2], 5400000);
    var oFill1 = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
    var oStroke = Api.CreateStroke(3 * 36000, oFill1);
    oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000);
}
```