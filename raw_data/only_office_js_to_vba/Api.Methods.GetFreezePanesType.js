# Freeze First Column and Display Freeze Type / Зафиксировать первый столбец и отобразить тип заморозки

**English:**  
This code freezes the first column and inserts the freeze type into cells A1 and B1.

**Russian:**  
Этот код фиксирует первый столбец и вставляет тип заморозки в ячейки A1 и B1.

```vba
' VBA code to freeze the first column and display the freeze type

Sub FreezeFirstColumn()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Freeze the first column
    With ws
        .Activate
        .Range("B1").Select
        ActiveWindow.FreezePanes = True
    End With
    
    ' Set values in A1 and B1
    ws.Range("A1").Value = "Type: "
    ws.Range("B1").Value = "column" ' VBA does not have a direct method to get freeze type
End Sub
```

```javascript
// JavaScript code to freeze the first column and display the freeze type

// This example freezes the first column and inserts the freeze type into the table.
Api.SetFreezePanesType('column');
var oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("A1").SetValue("Type: ");
oWorksheet.GetRange("B1").SetValue(Api.GetFreezePanesType());
```