`This code freezes the top row in an Excel spreadsheet.
Этот код фиксирует верхнюю строку в таблице Excel.`

```vba
' VBA code to freeze the top row in the active worksheet

Sub FreezeTopRow()
    ' Set the split row to the first row
    ActiveWindow.SplitRow = 1
    ' Enable freeze panes to freeze the specified rows
    ActiveWindow.FreezePanes = True
End Sub
```

```javascript
// JavaScript code to freeze the top row in OnlyOffice

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();
// Get the freeze panes object for the worksheet
var oFreezePanes = oWorksheet.GetFreezePanes();
// Freeze the first row in the worksheet
oFreezePanes.FreezeRows(1);
```