**Description / Описание**

*English: This code freezes the first column in the worksheet and inserts the address of the frozen range into cells A1 and B1.*

*Russian: Этот код замораживает первый столбец на листе и вставляет адрес замороженного диапазона в ячейки A1 и B1.*

```vba
' VBA code to freeze the first column and display the freeze panes location

Sub FreezeFirstColumn()
    Dim ws As Worksheet
    Dim freezeAddress As String
    
    Set ws = ActiveSheet
    
    ' Freeze the first column
    With ActiveWindow
        .SplitColumn = 1 ' Split before the second column
        .FreezePanes = True
    End With
    
    ' Get the freeze panes location
    freezeAddress = ws.Cells(1, 2).Address ' Cell B1 is the split location
    ws.Range("A1").Value = "Location: "
    ws.Range("B1").Value = freezeAddress
End Sub
```

```javascript
// This code freezes the first column and inserts the address of the frozen range into cells A1 and B1.

Api.SetFreezePanesType('column'); // Freeze the first column
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFreezePanes = oWorksheet.GetFreezePanes(); // Get the freeze panes
var oRange = oFreezePanes.GetLocation(); // Get the location of the freeze panes
oWorksheet.GetRange("A1").SetValue("Location: "); // Set "Location: " in A1
oWorksheet.GetRange("B1").SetValue(oRange.GetAddress()); // Set the address in B1
```