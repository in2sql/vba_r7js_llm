### This code demonstrates how to subscribe to the "onWorksheetChange" event, set a value in cell A1, and log changes.
### Этот код демонстрирует, как подписаться на событие "onWorksheetChange", установить значение в ячейке A1 и записать изменения.

```javascript
// JavaScript OnlyOffice API Code
// This code subscribes to the "onWorksheetChange" event,
// sets the value of cell A1 to "1",
// and logs the address of the changed range.

var oWorksheet = Api.GetActiveSheet();
var oRange = oWorksheet.GetRange("A1");
oRange.SetValue("1");
Api.attachEvent("onWorksheetChange", function(oRange){
    console.log("onWorksheetChange");
    console.log(oRange.GetAddress());
});
```

```vba
' VBA Code Equivalent
' This code subscribes to the WorksheetChange event,
' sets the value of cell A1 to "1",
' and logs the address of the changed range.

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Debug.Print "onWorksheetChange"
    Debug.Print Target.Address
End Sub

Sub SetCellValue()
    ' Sets the value of cell A1 to "1"
    ThisWorkbook.ActiveSheet.Range("A1").Value = "1"
End Sub
```