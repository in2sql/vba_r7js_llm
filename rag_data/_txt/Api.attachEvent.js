**Description / Описание:**

This code demonstrates how to subscribe to the "onWorksheetChange" event, set the value of cell A1 to "1", and log the change event with the affected cell's address.

Этот код демонстрирует, как подписаться на событие "onWorksheetChange", установить значение ячейки A1 равным "1" и записать в консоль информацию о событии изменения с адресом затронутой ячейки.

```vba
' VBA Code equivalent to subscribe to worksheet change event and set cell value

Private WithEvents ws As Worksheet

Private Sub Workbook_Open()
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Adjust the sheet name as needed
    ws.Range("A1").Value = "1"
End Sub

Private Sub ws_Change(ByVal Target As Range)
    If Not Intersect(Target, ws.Range("A1")) Is Nothing Then
        Debug.Print "onWorksheetChange"
        Debug.Print Target.Address
    End If
End Sub
```

```javascript
// JavaScript Code using OnlyOffice API to subscribe to worksheet change event and set cell value

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get the range A1 and set its value to "1"
var oRange = oWorksheet.GetRange("A1");
oRange.SetValue("1");

// Attach event handler for worksheet change
Api.attachEvent("onWorksheetChange", function(oRange){
    console.log("onWorksheetChange");
    console.log(oRange.GetAddress());
});
```