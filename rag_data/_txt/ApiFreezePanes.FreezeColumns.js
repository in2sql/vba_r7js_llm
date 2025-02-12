**Description / Описание**

This example freezes the first column.
Этот пример замораживает первый столбец.

```vba
' This VBA code freezes the first column in the active worksheet
Sub FreezeFirstColumn()
    With ActiveWindow
        .SplitColumn = 1       ' Set the split before the first column
        .FreezePanes = True    ' Enable freezing panes
    End With
End Sub
```

```javascript
// This OnlyOffice JS code freezes the first column
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFreezePanes = oWorksheet.GetFreezePanes(); // Get FreezePanes object
oFreezePanes.FreezeColumns(1); // Freeze the first column
```