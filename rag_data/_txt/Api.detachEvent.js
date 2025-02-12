# Description
This code sets the value of cell A1 to "1", attaches an event handler to the "Worksheet_Change" event to log changes, and then detaches the event handler.
Этот код устанавливает значение ячейки A1 в "1", привязывает обработчик события к событию "Worksheet_Change" для регистрации изменений, а затем отвязывает обработчик события.

```vba
' VBA Code
' This VBA code sets cell A1 to "1", attaches a Worksheet_Change event handler,
' logs the address of the changed range, and then detaches the event handler.

' Module: Standard Module
Public bEventAttached As Boolean

Sub SetValueAndAttachEvent()
    ' Set the value of cell A1 to "1"
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    oWorksheet.Range("A1").Value = "1"
    
    ' Attach event handler by setting a flag
    bEventAttached = True
End Sub

Sub DetachEvent()
    ' Detach event handler by clearing the flag
    bEventAttached = False
End Sub

' Module: Worksheet Module
Private Sub Worksheet_Change(ByVal Target As Range)
    If bEventAttached Then
        ' Log the address of the changed range
        Debug.Print "onWorksheetChange"
        Debug.Print Target.Address
    End If
End Sub
```

```javascript
// JavaScript Code
// This JS code sets cell A1 to "1", attaches an event listener to the onWorksheetChange event,
// logs the address of the changed range, and then detaches the event listener.

// Get active worksheet and set A1 to "1"
var oWorksheet = Api.GetActiveSheet();
var oRange = oWorksheet.GetRange("A1");
oRange.SetValue("1");

// Attach event handler for worksheet changes
Api.attachEvent("onWorksheetChange", function(oRange){
    console.log("onWorksheetChange");
    console.log(oRange.GetAddress());
});

// Detach event handler
Api.detachEvent("onWorksheetChange");
```