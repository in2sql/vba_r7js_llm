**Description:**
Freezes the first column in the active sheet.
  
Замораживает первый столбец на активном листе.

```vba
' VBA code to freeze the first column
Sub FreezeFirstColumn()
    With ActiveWindow
        .SplitColumn = 1
        .FreezePanes = True
    End With
End Sub
```

```javascript
// JavaScript code to freeze the first column
var oWorksheet = Api.GetActiveSheet();
// Get the freeze panes object
var oFreezePanes = oWorksheet.GetFreezePanes();
// Freeze the first column
oFreezePanes.FreezeColumns(1);
```