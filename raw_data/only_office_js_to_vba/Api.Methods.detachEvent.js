# Description / Описание

This code unsubscribes from the "onWorksheetChange" event, sets the value of cell A1 to "1", attaches an event handler to monitor worksheet changes, and then detaches the event handler.

Этот код отписывается от события "onWorksheetChange", устанавливает значение ячейки A1 на "1", прикрепляет обработчик событий для отслеживания изменений в листе и затем отвязывает обработчик событий.

## JavaScript Code / JavaScript код

```javascript
// This example unsubscribes from the "onWorksheetChange" event.

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get the range A1
var oRange = oWorksheet.GetRange("A1");

// Set the value of A1 to "1"
oRange.SetValue("1");

// Attach event handler to "onWorksheetChange"
Api.attachEvent("onWorksheetChange", function(oRange){
    console.log("onWorksheetChange");
    console.log(oRange.GetAddress());
});

// Detach the event handler from "onWorksheetChange"
Api.detachEvent("onWorksheetChange");
```

## VBA Code / VBA код

```vba
' This example unsubscribes from the "onWorksheetChange" event.

Dim oWorksheet As Worksheet
Dim oRange As Range

' Get the active worksheet
Set oWorksheet = ThisWorkbook.ActiveSheet

' Get the range A1
Set oRange = oWorksheet.Range("A1")

' Set the value of A1 to "1"
oRange.Value = "1"

' Attach event handler to "onWorksheetChange"
' Note: In VBA, event handlers are typically defined within the worksheet's code module.
' To dynamically attach and detach event handlers, a class module is required.
' Below is an example using a class module named "EventClass".

' --- In a Class Module named EventClass ---
' Public WithEvents Worksheet As Worksheet

' Private Sub Worksheet_Change(ByVal Target As Range)
'     If Not Intersect(Target, Worksheet.Range("A1")) Is Nothing Then
'         Debug.Print "onWorksheetChange"
'         Debug.Print Target.Address
'     End If
' End Sub

' --- In a Standard Module ---
' Dim EventHandler As New EventClass

' Sub AttachEvent()
'     Set EventHandler.Worksheet = ThisWorkbook.ActiveSheet
' End Sub

' Sub DetachEvent()
'     Set EventHandler.Worksheet = Nothing
' End Sub

' To use:
' Call AttachEvent  ' To attach the event handler
' Call DetachEvent  ' To detach the event handler
```