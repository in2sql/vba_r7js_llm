# Description / Описание

**English:** This code attaches an event handler to detect worksheet changes and sets the value of cell A1 to "1".

**Russian:** Этот код привязывает обработчик события для обнаружения изменений на листе и устанавливает значение ячейки A1 на "1".

## JavaScript Code

```javascript
// Attach event handler for worksheet changes
Api.attachEvent("onWorksheetChange", function(oRange){
    console.log("onWorksheetChange");
    console.log(oRange.GetAddress());
});

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get the range A1
var oRange = oWorksheet.GetRange("A1");

// Set the value of A1 to "1"
oRange.SetValue("1");
```

## VBA Code

```vba
' Attach event handler for worksheet changes
Private Sub Worksheet_Change(ByVal Target As Range)
    Debug.Print "onWorksheetChange"
    Debug.Print Target.Address
End Sub

' Set the value of A1 to "1"
Sub SetA1Value()
    ' Get the active worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set the value of A1 to "1"
    ws.Range("A1").Value = "1"
End Sub
```