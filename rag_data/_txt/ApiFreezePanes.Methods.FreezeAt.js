# Freeze a specified range in the worksheet / Заморозить указанную область в листе

```vba
' VBA code to freeze panes at the specified range

Sub FreezePanes()
    With ActiveWindow
        .SplitColumn = 7 ' Column H
        .SplitRow = 1    ' Row 2
        .FreezePanes = True
    End With
End Sub
```

```javascript
// JavaScript code to freeze panes at the specified range

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFreezePanes = oWorksheet.GetFreezePanes(); // Get the freeze panes object
var oRange = Api.GetRange('H2:K4'); // Define the range to freeze
oFreezePanes.FreezeAt(oRange); // Freeze panes at the specified range
```