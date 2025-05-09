# Description / Описание

**English:**  
This code gets the active worksheet and sets the value of the selected range to "selected".

**Russian:**  
Этот код получает активный лист и устанавливает значение выбранного диапазона на "selected".

```vba
' This code gets the active worksheet and sets the value of the selected range to "selected".

Sub SetSelectedRangeValue()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set the value of the selected range to "selected"
    Selection.Value = "selected"
End Sub
```

```javascript
// This code gets the active worksheet and sets the value of the selected range to "selected".

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
Api.GetSelection().SetValue("selected"); // Set the value of the selected range to "selected"
```