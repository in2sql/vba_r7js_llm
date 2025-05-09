**Description:**
This code freezes the top row of the active worksheet.
Этот код замораживает верхнюю строку активного листа.

```vba
' VBA code to freeze the top row of the active worksheet
Sub FreezeTopRow()
    ' Disable any existing freeze panes
    ActiveWindow.FreezePanes = False
    
    ' Select the second row
    Rows("2:2").Select
    
    ' Freeze panes above the selected row
    ActiveWindow.FreezePanes = True
End Sub
```

```javascript
// This code freezes the top row of the active worksheet
var oWorksheet = Api.GetActiveSheet();
var oFreezePanes = oWorksheet.GetFreezePanes();
oFreezePanes.FreezeRows(1);
```