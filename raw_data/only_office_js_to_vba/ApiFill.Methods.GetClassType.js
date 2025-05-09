```javascript
// This code gets a class type and inserts it into the document.
// Этот код получает тип класса и вставляет его в документ.

// JavaScript Code
// OnlyOffice API code
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oGs1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0); // Create first gradient stop
var oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000); // Create second gradient stop
var oFill = Api.CreateLinearGradientFill([oGs1, oGs2], 5400000); // Create linear gradient fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create stroke with no fill
oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000); // Add shape to worksheet
var sClassType = oFill.GetClassType(); // Get class type from fill
oWorksheet.SetColumnWidth(0, 15); // Set width for column A
oWorksheet.SetColumnWidth(1, 10); // Set width for column B
oWorksheet.GetRange("A1").SetValue("Class Type = " + sClassType); // Set value in cell A1
```

```vba
' This code gets a class type and inserts it into the document.
' Этот код получает тип класса и вставляет его в документ.

' Excel VBA Code
Sub InsertClassType()
    ' Get the active worksheet
    Dim oWorksheet As Object
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Create gradient stops
    Dim oGs1 As Object
    Set oGs1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0)
    
    Dim oGs2 As Object
    Set oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000)
    
    ' Create linear gradient fill
    Dim oFill As Object
    Set oFill = Api.CreateLinearGradientFill(Array(oGs1, oGs2), 5400000)
    
    ' Create stroke with no fill
    Dim oStroke As Object
    Set oStroke = Api.CreateStroke(0, Api.CreateNoFill())
    
    ' Add shape to worksheet
    oWorksheet.AddShape "flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000
    
    ' Get class type from fill
    Dim sClassType As String
    sClassType = oFill.GetClassType()
    
    ' Set column widths
    oWorksheet.SetColumnWidth 0, 15 ' Column A
    oWorksheet.SetColumnWidth 1, 10 ' Column B
    
    ' Set value in cell A1
    oWorksheet.GetRange("A1").SetValue "Class Type = " & sClassType
End Sub
```