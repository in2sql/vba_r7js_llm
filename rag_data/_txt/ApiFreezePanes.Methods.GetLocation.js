**Description / Описание**

This script freezes the first column in the active worksheet and pastes the address of the frozen range into cells A1 and B1.

Этот скрипт замораживает первый столбец в активном листе и вставляет адрес замороженного диапазона в ячейки A1 и B1.

---

```vba
' VBA Code to freeze the first column and insert the frozen range address

Sub FreezeFirstColumnAndInsertAddress()
    Dim ws As Worksheet
    Dim freezeRange As Range
    Dim freezeAddress As String

    ' Set the active worksheet
    Set ws = ActiveSheet

    ' Freeze the first column
    ws.Activate
    ws.Columns("B").Select
    ActiveWindow.FreezePanes = True

    ' Get the address of the frozen panes
    Set freezeRange = ws.Range("A1")
    freezeAddress = freezeRange.Address

    ' Insert values into cells A1 and B1
    ws.Range("A1").Value = "Location: "
    ws.Range("B1").Value = freezeAddress
End Sub
```

```javascript
// JavaScript Code to freeze the first column and insert the frozen range address

// Freeze the first column
Api.SetFreezePanesType('column');

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get the freeze panes object
var oFreezePanes = oWorksheet.GetFreezePanes();

// Get the location of the frozen panes
var oRange = oFreezePanes.GetLocation();

// Insert values into cells A1 and B1
oWorksheet.GetRange("A1").SetValue("Location: ");
oWorksheet.GetRange("B1").SetValue(oRange.GetAddress());
```