### Get the Active Sheet and Set Selected Range Value
Получение активного листа и установка значения выбранного диапазона.

```vba
' Get the Active Sheet and Set Selected Range Value

Dim oWorksheet As Worksheet
Set oWorksheet = ActiveSheet ' Get the active worksheet
Selection.Value = "selected" ' Set the value of the selected range
```

```javascript
// Get the Active Sheet and Set Selected Range Value
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
Api.GetSelection().SetValue("selected"); // Set the value of the selected range
```