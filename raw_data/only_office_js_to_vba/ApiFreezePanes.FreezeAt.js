**Description / Описание:**

This code freezes the specified range in the top-and-left-most pane of the worksheet.  
Этот код замораживает указанный диапазон в верхней и левой части окна рабочего листа.

---

**VBA Code:**

```vba
' This code freezes the specified range in the top-and-left-most pane of the worksheet.

Sub FreezePaneExample()
    ' Set the active sheet
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Freeze panes to the left of column H and above row 2
    With ActiveWindow
        .SplitColumn = 7 ' Column G is 7, so column H is 8
        .SplitRow = 1    ' Freeze above row 2
        .FreezePanes = True
    End With
End Sub
```

---

**OnlyOffice JS Code:**

```javascript
// This example freezes the specified range in top-and-left-most pane of the worksheet.

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFreezePanes = oWorksheet.GetFreezePanes(); // Get the FreezePanes object
var oRange = Api.GetRange('H2:K4'); // Define the range to freeze
oFreezePanes.FreezeAt(oRange); // Apply the freeze at the specified range
```