# Worksheet Change Event Handling / Обработка события изменения листа

**English:**  
This code sets up an event handler for worksheet changes, logs the address of the changed range, and sets the value of cell A1 to "1".

**Russian:**  
Этот код настраивает обработчик события изменения листа, выводит адрес измененного диапазона и устанавливает значение ячейки A1 равным "1".

```vba
' VBA Code

' This VBA code should be placed in the Worksheet module (e.g., Sheet1)

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Log the change event
    Debug.Print "onWorksheetChange"
    Debug.Print Target.Address
End Sub

Sub SetValueA1()
    ' Get the active sheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Get the range A1
    Dim oRange As Range
    Set oRange = oWorksheet.Range("A1")
    
    ' Set the value of A1 to "1"
    oRange.Value = "1"
End Sub
```

```javascript
// JavaScript Code

// Attach event handler for worksheet change
Api.attachEvent("onWorksheetChange", function(oRange){
    // Log the change event
    console.log("onWorksheetChange");
    console.log(oRange.GetAddress());
});

// Get the active sheet
var oWorksheet = Api.GetActiveSheet();

// Get the range A1
var oRange = oWorksheet.GetRange("A1");

// Set the value of A1 to "1"
oRange.SetValue("1");
```