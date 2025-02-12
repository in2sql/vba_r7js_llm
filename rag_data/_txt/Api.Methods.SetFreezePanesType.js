**Description / Описание:**

This code freezes the first column and inserts the address of the frozen range into the table.  
Этот код фиксирует первый столбец и вставляет адрес зафиксированного диапазона в таблицу.

```vba
' This VBA code freezes the first column and inserts the address of the frozen range into the table
Sub FreezeFirstColumn()
    Dim ws As Worksheet
    Dim freezeRange As Range
    Dim address As String
    
    Set ws = ActiveSheet
    
    ' Freeze the first column by selecting the second column and applying FreezePanes
    ws.Columns("B").Select
    ActiveWindow.FreezePanes = True
    
    ' Get the address of the frozen panes (first column)
    Set freezeRange = ws.Range("A1")
    address = freezeRange.Address
    
    ' Insert the location address into cells A1 and B1
    ws.Range("A1").Value = "Location: "
    ws.Range("B1").Value = address
End Sub
```

```javascript
// This JavaScript code freezes the first column and inserts the address of the frozen range into the table
Api.SetFreezePanesType('column');
var oWorksheet = Api.GetActiveSheet();
var oFreezePanes = oWorksheet.GetFreezePanes();
var oRange = oFreezePanes.GetLocation();
oWorksheet.GetRange("A1").SetValue("Location: ");
oWorksheet.GetRange("B1").SetValue(oRange.GetAddress());
```