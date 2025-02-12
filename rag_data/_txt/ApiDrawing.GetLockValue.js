# Description / Описание

**English:**  
This code demonstrates how to add a shape to the active sheet in OnlyOffice, set its properties, lock its selection, and display whether the shape is locked from selection in cell A1.

**Russian:**  
Этот код демонстрирует, как добавить фигуру на активный лист в OnlyOffice, установить её свойства, заблокировать её выбор и отобразить, заблокирована ли фигура для выбора в ячейке A1.

```vba
' VBA code to add a shape, set properties, lock selection, and display lock status
Sub AddLockedShape()
    ' Get the active worksheet
    Dim oWorksheet As Object
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Create a solid fill with RGB color (255, 111, 61)
    Dim oFill As Object
    Set oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61))
    
    ' Create a stroke with no fill
    Dim oStroke As Object
    Set oStroke = Api.CreateStroke(0, Api.CreateNoFill())
    
    ' Add a shape to the worksheet
    Dim oDrawing As Object
    Set oDrawing = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000)
    
    ' Set the size of the shape
    oDrawing.SetSize 120 * 36000, 70 * 36000
    
    ' Set the position of the shape
    oDrawing.SetPosition 0, 2 * 36000, 1, 3 * 36000
    
    ' Lock the shape from being selected
    oDrawing.SetLockValue "noSelect", True
    
    ' Get the lock status
    Dim bLockValue As Boolean
    bLockValue = oDrawing.GetLockValue("noSelect")
    
    ' Display the lock status in cell A1
    oWorksheet.GetRange("A1").SetValue "This drawing cannot be selected: " & bLockValue
End Sub
```

```javascript
// JavaScript code to add a shape, set properties, lock selection, and display lock status
function addLockedShape() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Create a solid fill with RGB color (255, 111, 61)
    var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
    
    // Create a stroke with no fill
    var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    
    // Add a shape to the worksheet
    var oDrawing = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
    
    // Set the size of the shape
    oDrawing.SetSize(120 * 36000, 70 * 36000);
    
    // Set the position of the shape
    oDrawing.SetPosition(0, 2 * 36000, 1, 3 * 36000);
    
    // Lock the shape from being selected
    oDrawing.SetLockValue("noSelect", true);
    
    // Get the lock status
    var bLockValue = oDrawing.GetLockValue("noSelect");
    
    // Display the lock status in cell A1
    oWorksheet.GetRange("A1").SetValue("This drawing cannot be selected: " + bLockValue); 
}
```